import json
import queue
import threading
import tkinter as tk
import webbrowser
from datetime import datetime, timedelta
from tkinter import ttk, messagebox, simpledialog
from .core import Auth, Graph, addresses, ai_write, config, save_config, secret, ident


class App:
    def __init__(self, args):
        self.root = tk.Tk()
        self.root.title('Outlook Workspace')
        self.root.geometry('960x740')
        self.root.minsize(720, 580)
        self.root.configure(bg='#eef2f7')
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure('TButton', padding=8)
        style.configure('TLabel', padding=3)
        self.frame = ttk.Frame(self.root, padding=22)
        self.frame.pack(fill='both', expand=True)
        self.status = tk.StringVar(value='Ready')
        ttk.Label(self.root, textvariable=self.status).pack(fill='x', padx=22, pady=8)
        self.events = queue.Queue()
        self.busy = False
        self.locked_widgets = []
        self.generation = 0
        self.root.protocol('WM_DELETE_WINDOW', self.close)
        self.root.after(100, self.poll)
        self.args = args
        self.show(args[0] if args else 'settings', *(args[1:]))

    def run(self):
        self.root.mainloop()

    def poll(self):
        try:
            while True:
                callback, value, error, generation = self.events.get_nowait()
                self.busy = False
                for widget, state in self.locked_widgets:
                    if widget.winfo_exists():
                        widget.configure(state=state)
                self.locked_widgets = []
                self.status.set('Ready' if not error else 'Action failed')
                if generation != self.generation:
                    continue
                if error:
                    messagebox.showerror('Outlook Workspace', str(error), parent=self.root)
                else:
                    callback(value)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def work(self, task, callback=lambda _: None):
        if self.busy:
            messagebox.showinfo('Please wait', 'An operation is still running.', parent=self.root)
            return
        self.busy = True
        def lock(parent):
            for widget in parent.winfo_children():
                lock(widget)
                if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Combobox, ttk.Checkbutton, tk.Text)):
                    state = str(widget.cget('state'))
                    self.locked_widgets.append((widget, state))
                    widget.configure(state='disabled')
        lock(self.frame)
        self.status.set('Working…')
        generation = self.generation
        def worker():
            try:
                value = task()
                self.events.put((callback, value, None, generation))
            except Exception as exc:
                self.events.put((callback, None, exc, generation))
        threading.Thread(target=worker, daemon=True).start()

    def close(self):
        if self.busy:
            messagebox.showinfo('Please wait', 'Wait for the current operation to finish before closing.', parent=self.root)
            return
        if getattr(self, 'dirty_editor', False) and not messagebox.askyesno('Close editor?', 'Close this editor? Unsaved changes will be lost.', parent=self.root):
            return
        self.root.destroy()

    def clear(self, title):
        self.generation += 1
        self.dirty_editor = False
        for widget in self.frame.winfo_children():
            widget.destroy()
        ttk.Label(self.frame, text=title, font=('Segoe UI', 20, 'bold')).pack(anchor='w', pady=(0, 14))
        nav = ttk.Frame(self.frame)
        nav.pack(fill='x', pady=(0, 10))
        for label, screen in [('Mail', 'mail'), ('Compose', 'compose'), ('Drafts', 'drafts'), ('Calendar', 'calendar'), ('To Do', 'tasks'), ('Settings', 'settings')]:
            ttk.Button(nav, text=label, command=lambda s=screen: self.navigate(s)).pack(side='left', padx=2)

    def navigate(self, screen):
        if self.busy:
            return
        if getattr(self, 'dirty_editor', False) and not messagebox.askyesno('Leave editor?', 'Leave this editor? Unsaved changes will be lost.', parent=self.root):
            return
        self.dirty_editor = False
        self.show(screen)

    def show(self, screen, *args):
        if screen == 'settings':
            self.settings()
        elif screen == 'compose':
            self.compose(prompt=args[0] if args else '')
        elif screen == 'message':
            self.clear('Loading email')
            self.work(lambda: Graph().message(args[0]), lambda m: self.message(m, args[1] if len(args) > 1 else ''))
        elif screen == 'event':
            if args:
                self.clear('Loading event')
                self.work(lambda: Graph().request('GET', '/me/events/' + ident(args[0])), self.event)
            else:
                self.event({})
        elif screen == 'calendar':
            self.calendar()
        elif screen == 'tasks':
            self.tasks(args[0] if args else None)
        else:
            self.mail('drafts' if screen == 'drafts' else 'inbox')

    def field(self, label, value='', secret_value=False):
        ttk.Label(self.frame, text=label).pack(anchor='w')
        entry = ttk.Entry(self.frame, show='•' if secret_value else '')
        entry.insert(0, value)
        entry.pack(fill='x', pady=(0, 5))
        return entry

    def text(self, value='', height=12):
        frame = ttk.Frame(self.frame)
        frame.pack(fill='both', expand=True, pady=8)
        text = tk.Text(frame, height=height, wrap='word', font=('Segoe UI', 11), undo=True)
        scroll = ttk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        text.pack(fill='both', expand=True)
        text.insert('1.0', value)
        return text

    def buttons(self, actions):
        bar = ttk.Frame(self.frame)
        bar.pack(fill='x', pady=6)
        for label, callback in actions:
            ttk.Button(bar, text=label, command=callback).pack(side='left', padx=3)

    def settings(self):
        self.clear('Settings & connection')
        values = config()
        cid = self.field('Microsoft Entra application (client) ID', values.get('client_id', ''))
        tenant = self.field('Tenant: common, organizations, consumers, or your tenant ID', values.get('tenant', 'common'))
        model = self.field('OpenAI model', values.get('model', 'gpt-4.1-mini'))
        key = self.field('OpenAI API key — leave blank to retain the saved key', secret_value=True)
        ttk.Label(self.frame, text='Register a public desktop app with http://localhost redirect. No client secret.\nAI writing sends the draft and selected reference email to OpenAI only when you click Generate.', wraplength=830).pack(anchor='w', pady=12)
        def save():
            save_config({'client_id': cid.get().strip(), 'tenant': tenant.get().strip() or 'common', 'model': model.get().strip() or 'gpt-4.1-mini'})
            if key.get().strip():
                secret('openai-key', key.get().strip())
                key.delete(0, 'end')
        def login():
            try:
                save()
            except Exception as exc:
                messagebox.showerror('Settings', str(exc))
                return
            self.work(lambda: Auth().login(), lambda name: self.status.set('Signed in: ' + name))
        self.buttons([('Save settings', lambda: self.work(save, lambda _: self.status.set('Settings saved'))), ('Sign in with Microsoft', login), ('Sign out', lambda: self.work(lambda: Auth().logout())), ('Remove AI key', lambda: self.work(lambda: secret('openai-key', '')))])
        self.buttons([('Microsoft app registration', lambda: webbrowser.open('https://entra.microsoft.com/')), ('OpenAI API keys', lambda: webbrowser.open('https://platform.openai.com/api-keys')), ('Open setup guide', self.guide)])

    def guide(self):
        from .core import ROOT
        import os
        if hasattr(os, 'startfile'):
            os.startfile(str(ROOT / 'README.md'))
        else:
            webbrowser.open((ROOT / 'README.md').as_uri())

    def mail(self, folder='inbox'):
        self.clear('Email · ' + folder)
        search = self.field('Search subject, sender or email content')
        folder_value = tk.StringVar(value=folder)
        ttk.Combobox(self.frame, textvariable=folder_value, values=['inbox', 'drafts', 'sentitems', 'archive', 'deleteditems'], state='readonly').pack(anchor='w')
        tree = ttk.Treeview(self.frame, columns=('sender', 'date'), show='tree headings')
        tree.heading('#0', text='Subject')
        tree.heading('sender', text='From')
        tree.heading('date', text='Received')
        tree.column('#0', width=380)
        tree.column('sender', width=200)
        tree.column('date', width=160)
        tree.pack(fill='both', expand=True, pady=12)
        records = {}
        def loaded(items):
            tree.delete(*tree.get_children())
            records.clear()
            for i, item in enumerate(items):
                records[str(i)] = item
                tree.insert('', 'end', iid=str(i), text=('● ' if not item.get('isRead') else '') + (item.get('subject') or '(No subject)'), values=(item.get('from', {}).get('emailAddress', {}).get('address', ''), item.get('receivedDateTime', '')))
            self.status.set(f'{len(items)} messages (up to 60 per search)')
        def refresh():
            f, term = folder_value.get(), search.get()
            self.work(lambda: Graph().messages(f, term), loaded)
        def open_selected(_=None):
            selected = tree.selection()
            if selected:
                self.show('message', records[selected[0]]['id'])
        self.buttons([('Search / refresh', refresh), ('Open selected', open_selected), ('New draft', lambda: self.compose())])
        tree.bind('<Double-1>', open_selected)
        search.bind('<Return>', lambda _: refresh())
        refresh()

    def message(self, item, action=''):
        if item.get('isDraft'):
            self.compose(item)
            return
        self.clear(item.get('subject') or '(No subject)')
        sender = item.get('from', {}).get('emailAddress', {}).get('address', '')
        ttk.Label(self.frame, text='From: ' + sender).pack(anchor='w')
        ttk.Label(self.frame, text='To: ' + '; '.join(x['emailAddress']['address'] for x in item.get('toRecipients', []))).pack(anchor='w')
        body = self.text(item.get('body', {}).get('content', ''))
        body.configure(state='disabled')
        def reply(all_recipients=False, forward=False):
            self.work(lambda: Graph().reply(item['id'], all_recipients, forward), lambda draft: self.compose(draft, context=item.get('body', {}).get('content', '')))
        def manage(kind):
            if kind in ('delete', 'archive') and not messagebox.askyesno('Confirm', 'Move this email to ' + ('Deleted Items' if kind == 'delete' else 'Archive') + '?'):
                return
            self.work(lambda: Graph().mail_action(item['id'], kind), lambda _: self.mail())
        self.buttons([('Reply', reply), ('Reply all', lambda: reply(True)), ('Forward', lambda: reply(forward=True)), ('Open in Outlook', lambda: webbrowser.open(item.get('webLink', 'https://outlook.office.com/mail/')))])
        self.buttons([('Mark unread' if item.get('isRead') else 'Mark read', lambda: manage('unread' if item.get('isRead') else 'read')), ('Flag', lambda: manage('flag')), ('Archive', lambda: manage('archive')), ('Delete', lambda: manage('delete'))])
        if action in ('reply', 'replyall', 'forward'):
            reply(action == 'replyall', action == 'forward')

    def compose(self, item=None, prompt='', context=''):
        item = item or {}
        self.clear('Draft editor')
        self.dirty_editor = True
        mid = [item.get('id')]
        def recipients(name):
            return '; '.join(x['emailAddress']['address'] for x in item.get(name, []))
        to = self.field('To', recipients('toRecipients'))
        cc = self.field('Cc', recipients('ccRecipients'))
        bcc = self.field('Bcc', recipients('bccRecipients'))
        subject = self.field('Subject', item.get('subject', ''))
        body = self.text(item.get('body', {}).get('content', ''), height=9)
        instruction = self.field('AI writing request (e.g. make concise, write a polite follow-up)', prompt)
        generated = [None]
        def generate():
            if not instruction.get().strip():
                messagebox.showinfo('AI writing', 'Enter a writing request first.')
                return
            text = body.get('1.0', 'end-1c')
            title, request = subject.get(), instruction.get()
            def preview(value):
                generated[0] = value
                window = tk.Toplevel(self.root)
                window.title('Review AI suggestion')
                window.geometry('760x520')
                editor = tk.Text(window, wrap='word', font=('Segoe UI', 11))
                editor.pack(fill='both', expand=True, padx=15, pady=15)
                editor.insert('1.0', value)
                def apply():
                    body.delete('1.0', 'end')
                    body.insert('1.0', editor.get('1.0', 'end-1c'))
                    window.destroy()
                ttk.Button(window, text='Apply to draft', command=apply).pack(side='left', padx=15, pady=10)
                ttk.Button(window, text='Discard', command=window.destroy).pack(side='left')
            self.work(lambda: ai_write(request, title, text, context), preview)
        def save(send=False):
            data = (subject.get(), body.get('1.0', 'end-1c'), to.get(), cc.get(), bcc.get())
            try:
                addresses(data[2]); addresses(data[3]); addresses(data[4])
                if send and not addresses(data[2]):
                    raise ValueError('Add a To recipient.')
            except ValueError as exc:
                messagebox.showerror('Recipients', str(exc)); return
            if send and not messagebox.askyesno('Send email?', f'Send “{data[0]}” to {data[2]}?\nCc: {data[3]}\nBcc: {data[4]}'):
                return
            def operation():
                graph = Graph()
                saved = graph.draft(*data, mid=mid[0])
                # Keep ID even if sending fails, preventing accidental duplicate drafts.
                mid[0] = saved['id']
                if send:
                    graph.send(mid[0])
                return saved
            def done(saved):
                mid[0] = saved['id']
                self.dirty_editor = False
                if send:
                    self.mail('sentitems')
                else:
                    # Future edits still need the leave-editor warning.
                    self.dirty_editor = True
                    self.status.set('Draft saved to Outlook')
            self.work(operation, done)
        self.buttons([('Generate with AI', generate), ('Save draft', save), ('Send…', lambda: save(True))])
        ttk.Label(self.frame, text='Generate sends this draft and reply context to OpenAI. Review the suggestion before applying.').pack(anchor='w')

    def calendar(self):
        self.clear('Calendar · next 7 days')
        tree = ttk.Treeview(self.frame, columns=('start', 'location'), show='tree headings')
        tree.heading('#0', text='Event'); tree.heading('start', text='Start (event timezone)'); tree.heading('location', text='Location')
        tree.pack(fill='both', expand=True)
        records = {}
        def loaded(items):
            tree.delete(*tree.get_children()); records.clear()
            for i, item in enumerate(items):
                records[str(i)] = item
                tree.insert('', 'end', iid=str(i), text=item.get('subject', ''), values=(item['start']['dateTime'] + ' ' + item['start'].get('timeZone', ''), item.get('location', {}).get('displayName', '')))
        def open_item(_=None):
            if tree.selection():
                self.event(records[tree.selection()[0]])
        self.buttons([('Refresh', lambda: self.work(lambda: Graph().agenda(), loaded)), ('New event', lambda: self.event({})), ('Edit selected', open_item)])
        tree.bind('<Double-1>', open_item)
        self.work(lambda: Graph().agenda(), loaded)

    def event(self, item):
        self.clear('Edit event' if item else 'New event')
        self.dirty_editor = True
        now = datetime.now().replace(second=0, microsecond=0)
        subject = self.field('Title', item.get('subject', ''))
        start = self.field('Start (YYYY-MM-DDTHH:MM:SS)', item.get('start', {}).get('dateTime', now.isoformat()))
        end = self.field('End (YYYY-MM-DDTHH:MM:SS)', item.get('end', {}).get('dateTime', (now + timedelta(hours=1)).isoformat()))
        zone = self.field('Microsoft timezone (e.g. UTC, US Mountain Standard Time)', item.get('start', {}).get('timeZone', 'UTC'))
        if not item:
            start.delete(0, 'end'); start.insert(0, datetime.utcnow().replace(second=0, microsecond=0).isoformat())
            end.delete(0, 'end'); end.insert(0, (datetime.utcnow().replace(second=0, microsecond=0) + timedelta(hours=1)).isoformat())
        location = self.field('Location', item.get('location', {}).get('displayName', ''))
        people = self.field('Attendees (email addresses)', '; '.join(x['emailAddress']['address'] for x in item.get('attendees', [])))
        online = tk.BooleanVar(value=item.get('isOnlineMeeting', False))
        ttk.Checkbutton(self.frame, text='Create Teams meeting (requires a supported Microsoft account)', variable=online).pack(anchor='w')
        body = self.text(item.get('body', {}).get('content', ''), 5)
        if item.get('isOnlineMeeting'):
            body.configure(state='disabled')
            ttk.Label(self.frame, text='Edit the description in Outlook to preserve the online meeting link.').pack(anchor='w')
        def save():
            try:
                data = {'subject': subject.get(), 'start': {'dateTime': start.get(), 'timeZone': zone.get()}, 'end': {'dateTime': end.get(), 'timeZone': zone.get()}, 'location': {'displayName': location.get()}, 'body': {'contentType': 'Text', 'content': body.get('1.0', 'end-1c')}, 'attendees': [dict(x, type='required') for x in addresses(people.get())]}
                if not data['subject'].strip():
                    raise ValueError('Enter an event title.')
                if item.get('isOnlineMeeting'):
                    # Graph requires preserving its embedded online meeting blob.
                    data.pop('body')
                old_attendees = {x['emailAddress']['address'].lower(): x for x in item.get('attendees', [])}
                data['attendees'] = [old_attendees.get(x['emailAddress']['address'].lower(), x) for x in data['attendees']]
                if online.get() and not item.get('isOnlineMeeting'):
                    data.update(isOnlineMeeting=True, onlineMeetingProvider='teamsForBusiness')
                datetime.fromisoformat(start.get()); datetime.fromisoformat(end.get())
            except ValueError as exc:
                messagebox.showerror('Event', str(exc)); return
            if people.get() and not messagebox.askyesno('Send invitations?', 'Saving this event sends invitations or updates to attendees. Continue?'):
                return
            self.work(lambda: Graph().event(data, item.get('id')), lambda _: self.saved_calendar())
        actions = [('Save event', save)]
        if item:
            def delete():
                if messagebox.askyesno('Delete event?', 'Delete this event? Attendees may receive cancellation notices.'):
                    self.work(lambda: Graph().delete_event(item['id']), lambda _: self.saved_calendar())
            actions.append(('Delete…', delete))
        self.buttons(actions)

    def saved_calendar(self):
        self.dirty_editor = False
        self.calendar()

    def tasks(self, selected=None):
        self.clear('Microsoft To Do')
        list_var = tk.StringVar()
        combo = ttk.Combobox(self.frame, textvariable=list_var, state='readonly')
        combo.pack(fill='x', pady=8)
        tree = ttk.Treeview(self.frame, columns=('status', 'due'), show='tree headings')
        tree.heading('#0', text='Task'); tree.heading('status', text='Status'); tree.heading('due', text='Due')
        tree.pack(fill='both', expand=True)
        lists, records = {}, {}
        def list_id():
            return lists.get(list_var.get())
        def loaded_tasks(items):
            tree.delete(*tree.get_children()); records.clear()
            for i, item in enumerate(items):
                records[str(i)] = item
                tree.insert('', 'end', iid=str(i), text=item['title'], values=(item.get('status'), item.get('dueDateTime', {}).get('dateTime', '')))
        def refresh(_=None):
            lid = list_id()
            if lid:
                self.work(lambda: Graph().tasks(lid), loaded_tasks)
        def loaded_lists(items):
            for item in items:
                lists[item['displayName'] + ' · ' + item['id'][-6:]] = item['id']
            combo['values'] = list(lists)
            if lists:
                list_var.set(next((k for k, v in lists.items() if v == selected), next(iter(lists))))
                refresh()
        def edit(_=None, new=False):
            lid = list_id()
            if not lid:
                return
            item = {} if new or not tree.selection() else records[tree.selection()[0]]
            if not new and not item:
                return
            self.task_editor(lid, item)
        def complete():
            if tree.selection():
                item = records[tree.selection()[0]]
                status = 'notStarted' if item.get('status') == 'completed' else 'completed'
                lid = list_id()
                self.work(lambda: Graph().task(lid, {'status': status}, item['id']), lambda _: refresh())
        def new_list():
            name = simpledialog.askstring('New list', 'List name:', parent=self.root)
            if name and name.strip():
                self.work(lambda: Graph().request('POST', '/me/todo/lists', {'displayName': name.strip()}), lambda _: self.tasks())
        self.buttons([('Refresh', refresh), ('New list', new_list), ('New task', lambda: edit(new=True)), ('Edit task', edit), ('Complete / reopen', complete)])
        combo.bind('<<ComboboxSelected>>', refresh)
        tree.bind('<Double-1>', edit)
        self.work(lambda: Graph().lists(), loaded_lists)

    def task_editor(self, lid, item):
        self.clear('Edit task' if item else 'New task')
        self.dirty_editor = True
        title = self.field('Title', item.get('title', ''))
        due = self.field('Due date (YYYY-MM-DD, optional)', item.get('dueDateTime', {}).get('dateTime', '')[:10])
        importance = tk.StringVar(value=item.get('importance', 'normal'))
        ttk.Label(self.frame, text='Importance').pack(anchor='w')
        ttk.Combobox(self.frame, textvariable=importance, values=['low', 'normal', 'high'], state='readonly').pack(anchor='w')
        complete = tk.BooleanVar(value=item.get('status') == 'completed')
        ttk.Checkbutton(self.frame, text='Completed', variable=complete).pack(anchor='w')
        body = self.text(item.get('body', {}).get('content', ''), 8)
        def done(_):
            self.dirty_editor = False
            self.tasks(lid)
        def save():
            data = {'title': title.get().strip(), 'importance': importance.get(), 'status': 'completed' if complete.get() else 'notStarted', 'body': {'contentType': 'text', 'content': body.get('1.0', 'end-1c')}, 'dueDateTime': None}
            try:
                if not data['title']:
                    raise ValueError('Enter a task title.')
                if due.get():
                    parsed = datetime.strptime(due.get(), '%Y-%m-%d')
                    data['dueDateTime'] = {'dateTime': parsed.isoformat(), 'timeZone': 'UTC'}
            except ValueError as exc:
                messagebox.showerror('Task', str(exc)); return
            self.work(lambda: Graph().task(lid, data, item.get('id')), done)
        actions = [('Save task', save)]
        if item:
            def delete():
                if messagebox.askyesno('Delete task?', 'Permanently delete this task?'):
                    self.work(lambda: Graph().delete_task(lid, item['id']), done)
            actions.append(('Delete…', delete))
        self.buttons(actions)

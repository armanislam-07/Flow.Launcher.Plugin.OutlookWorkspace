import subprocess
import sys
import webbrowser
from .core import Graph, ROOT, config


def result(title, subtitle='', method='launch', parameters=None, data=None, score=0):
    row = {'Title': title, 'SubTitle': subtitle, 'IcoPath': 'Images/app.png', 'Score': score,
           'JsonRPCAction': {'method': method, 'parameters': parameters or [], 'dontHideAfterAction': False}}
    if data:
        row['ContextData'] = data
    return row


class Launcher:
    def __init__(self, graph_factory=Graph):
        self.graph_factory = graph_factory

    def launch(self, screen='settings', *args):
        flags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        subprocess.Popen([sys.executable, str(ROOT / 'ui.py'), screen, *args], cwd=str(ROOT), creationflags=flags,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True

    def open_url(self, url):
        if url.startswith(('https://outlook.', 'https://to-do.')):
            webbrowser.open(url)
        return True

    def query(self, query):
        command, _, term = query.strip().partition(' ')
        command = command.lower()
        if not command:
            return [result('Search email', 'ou mail <search>', parameters=['mail']),
                    result('Compose email', 'Create a draft with optional AI writing', parameters=['compose']),
                    result('Drafts', 'Edit, rewrite and send saved drafts', parameters=['drafts']),
                    result('Calendar', 'View agenda and manage events', parameters=['calendar']),
                    result('Microsoft To Do', 'Manage lists, tasks, due dates and completion', parameters=['tasks']),
                    result('Settings and sign in', 'Microsoft account and OpenAI API key', parameters=['settings'])]
        if command in ('settings', 'login', 'logout', 'help'):
            return [result('Open settings and setup guide', parameters=['settings'])]
        if command in ('compose', 'new', 'ai'):
            return [result('Compose email', term or 'Write a new email', parameters=['compose', term])]
        if not config().get('client_id'):
            return [result('Set up Microsoft account', 'Enter your app client ID, then sign in', parameters=['settings'])]
        try:
            graph = self.graph_factory()
            if command in ('mail', 'inbox', 'drafts', 'sent'):
                folder = {'mail': 'inbox', 'sent': 'sentitems'}.get(command, command)
                rows = graph.messages(folder, term)
                return [result(x.get('subject') or '(No subject)',
                               ('Unread · ' if not x.get('isRead') else '') + x.get('from', {}).get('emailAddress', {}).get('address', '') + ' · ' + x.get('bodyPreview', '')[:160],
                               parameters=['message', x['id']], data={'kind': 'message', 'id': x['id']}) for x in rows] or [result('No messages found', 'Open mailbox', parameters=['mail'])]
            if command in ('calendar', 'agenda'):
                return [result('Create calendar event', parameters=['event'])] + [result(x.get('subject', '(No title)'),
                       x['start']['dateTime'] + ' ' + x['start'].get('timeZone', '') + ' · ' + x.get('location', {}).get('displayName', ''),
                       parameters=['event', x['id']], data={'kind': 'event', 'id': x['id']}) for x in graph.agenda(search=term)]
            if command in ('todo', 'tasks'):
                return [result('Open Microsoft To Do', 'Create lists and tasks', parameters=['tasks'])] + [result(x['displayName'], 'Open task list', parameters=['tasks', x['id']]) for x in graph.lists() if term.lower() in x['displayName'].lower()]
            return [result('Unknown command', 'Use mail, inbox, sent, drafts, compose, calendar, tasks, settings', parameters=['settings'])]
        except Exception as exc:
            return [result('Outlook needs attention', str(exc)[:220], parameters=['settings'])]

    def context_menu(self, data):
        if not isinstance(data, dict):
            return []
        if data.get('kind') == 'message':
            return [result(title, 'Open message and review action', parameters=['message', data['id'], action])
                    for title, action in [('Reply', 'reply'), ('Reply all', 'replyall'), ('Forward', 'forward'), ('Mark read / unread', 'read'), ('Archive or delete', 'manage')]]
        return []

# Outlook Workspace for Flow Launcher

A Windows Flow Launcher plugin for Microsoft Outlook / Exchange Online mail, Outlook Calendar, and Microsoft To Do. Launcher results open native desktop editors; Microsoft Graph is the data source. Personal Microsoft accounts and Microsoft 365 work/school accounts are supported, subject to organization policy and service availability. Local-only Outlook PST/POP accounts are not supported.

## Install

After a GitHub Release is published, install its `Flow.Launcher.Plugin.OutlookWorkspace.zip` with Flow's `pm install <release-asset-url>` command. Releases bundle Windows x64 dependencies for Python 3.11. For source installation or other Python versions, use the steps below. Repository/release/store submission instructions are in [PUBLISHING.md](PUBLISHING.md).

1. Install Windows Python 3.10+ with Tcl/Tk and pip. Use the same interpreter in Flow Launcher Settings → General → Python. Flow Launcher runs on Windows; macOS can run the unit tests but cannot host Flow Launcher.
2. Extract the distribution or clone this folder.
3. In PowerShell run `powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Python python`. For portable/custom Flow installations pass `-PluginRoot 'C:\path\to\FlowLauncher\Plugins'`.
4. Restart Flow Launcher. Type `ou settings` and press Enter.

The installer copies source and installs dependencies into the plugin's `lib` directory. Updating replaces plugin code and preserves account configuration in `%LOCALAPPDATA%\FlowOutlookWorkspace`. Do not copy the dependencies from a different OS or Python version.

## Microsoft account setup

Register your own application in [Microsoft Entra](https://entra.microsoft.com/) → App registrations → New registration.

- For both work/school and personal Outlook accounts choose “Accounts in any organizational directory and personal Microsoft accounts”. For organization-only deployments choose the appropriate tenant scope.
- Authentication → Add platform → Mobile and desktop applications → add `http://localhost` as a redirect URI. The interactive sign-in uses a random localhost port.
- API permissions → Microsoft Graph → **Delegated**: `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Tasks.ReadWrite`.
- No client secret is needed or supported. This is a public desktop client. Your organization may require administrator consent.
- Copy the Application (client) ID into plugin settings. Tenant `common` supports both account types; `organizations` restricts to work/school, `consumers` to personal accounts. A tenant ID restricts to your organization.
- Save settings, click **Sign in with Microsoft**, and complete browser consent. Sign out removes the local account token cache; it does not revoke the application's server-side consent.

Account tokens use Microsoft MSAL Extensions with OS encryption (Windows DPAPI) and a file lock shared across launcher/editor processes. The optional OpenAI key uses Windows Credential Manager through keyring. Settings contain only the client ID, tenant, and model. The plugin selects the first cached account; sign out before switching accounts. OS encryption and Windows Credential Manager are required; credentials are never saved to a plaintext fallback.

## Commands

| Command | Behavior |
| --- | --- |
| `ou` | Feature menu |
| `ou mail invoice` | Search Inbox with Microsoft Graph full-text search |
| `ou inbox` | Recent Inbox messages |
| `ou sent Alex` | Search Sent Items |
| `ou drafts` | Open saved drafts for editing / sending |
| `ou compose` | New draft editor |
| `ou ai write a polite follow-up` | Composer with a writing request prefilled |
| `ou calendar` | Next seven days, including recurring occurrences |
| `ou calendar planning` | Filter agenda titles |
| `ou tasks` / `ou todo` | Select a Microsoft To Do list |
| `ou settings` | Configuration, browser sign-in, sign-out |

Enter opens the selected result. Email context menus provide reply, reply-all, forward and management shortcuts. The mailbox editor also browses Inbox, Drafts, Sent Items, Archive, and Deleted Items. A search returns up to 60 messages; tasks and calendar lists return up to 200 items. Search results depend on Microsoft's mailbox indexing.

## Email and AI writing

Read mail, mark read/unread, flag, archive, move to Deleted Items, and create reply / reply-all / forward drafts. Draft editors support To, Cc, Bcc, subject and plain-text body. Save writes to Outlook Drafts; Send asks for confirmation with recipient details. Reply/forward creates a server-side draft immediately. Closing it leaves the draft available in Outlook.

For AI writing, enter an [OpenAI API key](https://platform.openai.com/api-keys) in Settings or set `OPENAI_API_KEY` in Flow Launcher's environment. Set the model to one available to your API account (default `gpt-4.1-mini`). API usage is separately billed; this plugin uses the OpenAI Responses API, not a logged-in ChatGPT browser session or ChatGPT subscription.

Enter a writing request and click **Generate with AI**. Only this explicit action sends your subject, current draft, and reply reference (when created from the current email viewer) to OpenAI. A separate preview lets you edit, apply or discard the suggestion. Generation never saves or sends email. `store=False` is set on Responses requests; your account's API data policies still apply. Received email is treated as untrusted context in the writing prompt.

## Calendar and To Do

Calendar supports the default calendar's next seven days, creating/editing/deleting events, timezone, location, attendees, and optional Teams meetings. Saving attendee changes asks for confirmation because Microsoft sends invitations/updates. Teams meeting creation requires a supporting account. When editing an existing online meeting, its existing meeting details are preserved. Editing a recurring occurrence changes that occurrence; there is no series editor.

To Do supports list creation/selection, task creation/editing/deletion, notes, importance, optional due date, completion and reopening. Clear a due date to remove it. Dates use UTC. Custom calendar selection, recurrence authoring, availability lookup, meeting responses, task recurrence/reminders/checklists, attachments, rich HTML editing, offline synchronization, and multi-account switching are not implemented in this version. Open a message in Outlook for attachments and full HTML rendering.

## Development and validation

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python -m compileall -q main.py ui.py outlook_flow scripts
python scripts/package.py
python ui.py settings
```

The package command creates the source-only `dist/Flow.Launcher.Plugin.OutlookWorkspace.zip`. The GitHub release workflow builds the installable Windows version with dependencies in `lib/`. Tests use mocked Microsoft/OpenAI services; no mail is sent and no cloud credentials are required. Before production use, validate on Windows with a test Microsoft account: consent, mailbox search, save/reopen draft, reply-all recipients, AI apply/discard, confirmed sending, calendar invitation and cancellation, task due dates and completion, and sign-out. Live Windows integration cannot be verified from macOS without those accounts.

## Troubleshooting

- No launcher results: ensure Flow uses the Python version used during install and restart it.
- “Open settings”: configure the client ID or sign in again.
- Authentication redirect errors: register `http://localhost` under Mobile and desktop applications.
- Graph 403: verify delegated permissions, consent, account service/license availability and tenant policy.
- Graph 429: wait for the displayed retry interval. Writes are not automatically replayed.
- AI error: verify API key, model access and API billing. The current draft is preserved.
- Native window does not open: run `python ui.py settings` from the installed plugin directory to see interpreter or Tcl/Tk errors.

API/protocol references: [Flow Python plugin guide](https://github.com/Flow-Launcher/docs/blob/main/py-write-code.md), [Microsoft Graph permissions](https://learn.microsoft.com/en-us/graph/permissions-reference), [draft replies](https://learn.microsoft.com/en-us/graph/api/message-createreply?view=graph-rest-1.0), [OpenAI Python SDK](https://developers.openai.com/api/reference/python).

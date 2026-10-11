"""Flow Launcher one-shot JSON-RPC entry point. Stdout is protocol only."""
import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'lib'))
from outlook_flow.launcher import Launcher


def dispatch(request, launcher=None):
    launcher = launcher or Launcher()
    methods = {'query': launcher.query, 'context_menu': launcher.context_menu,
               'launch': launcher.launch, 'open_url': launcher.open_url}
    method = request.get('method')
    if method not in methods:
        return {'jsonrpc': '2.0', 'id': request.get('id'), 'error': {'code': -32601, 'message': 'Unknown method'}}
    try:
        value = methods[method](*request.get('parameters', request.get('params', [])))
        return {'jsonrpc': '2.0', 'id': request.get('id'), 'result': value}
    except Exception:
        return {'jsonrpc': '2.0', 'id': request.get('id'), 'error': {'code': -32000, 'message': 'Unable to execute Outlook action. Open settings.'}}


if __name__ == '__main__':
    try:
        request = json.loads(sys.argv[1])
        print(json.dumps(dispatch(request), ensure_ascii=True))
    except Exception:
        print(json.dumps({'jsonrpc': '2.0', 'id': None, 'error': {'code': -32700, 'message': 'Invalid request'}}))

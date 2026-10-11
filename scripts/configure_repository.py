"""Set GitHub URLs in plugin metadata and regenerate the store submission."""
import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ASSET = 'Flow.Launcher.Plugin.OutlookWorkspace.zip'


def configure(repository, root=ROOT):
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+', repository):
        raise ValueError('Use GitHub owner/repository format.')
    path = root / 'plugin.json'
    plugin = json.loads(path.read_text('utf-8'))
    plugin['Website'] = f'https://github.com/{repository}'
    path.write_text(json.dumps(plugin, indent=2) + '\n', 'utf-8')
    submission = {key: plugin[key] for key in ['ID', 'Name', 'Description', 'Author', 'Version', 'Language', 'Website']}
    submission.update(
        UrlDownload=f'https://github.com/{repository}/releases/download/v{plugin["Version"]}/{ASSET}',
        UrlSourceCode=f'https://github.com/{repository}/tree/main',
        IcoPath=f'https://cdn.jsdelivr.net/gh/{repository}@v{plugin["Version"]}/Images/app.png')
    directory = root / 'submission'
    directory.mkdir(exist_ok=True)
    target = directory / f'OutlookWorkspace-{plugin["ID"]}.json'
    target.write_text(json.dumps(submission, indent=2) + '\n', 'utf-8')
    return target


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('repository')
    args = parser.parse_args()
    print(configure(args.repository))

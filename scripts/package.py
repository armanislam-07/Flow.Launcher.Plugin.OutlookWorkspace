"""Build a Flow installation ZIP, optionally bundling Windows dependencies."""
import argparse
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent.parent
ASSET = 'Flow.Launcher.Plugin.OutlookWorkspace.zip'


def package(root=ROOT, require_dependencies=False):
    manifest = json.loads((root / 'plugin.json').read_text('utf-8'))
    required = ['ID', 'ActionKeyword', 'Name', 'Description', 'Author', 'Version',
                'Language', 'Website', 'IcoPath', 'ExecuteFileName']
    for key in required:
        if not manifest.get(key):
            raise ValueError(f'Missing plugin metadata: {key}')
    if require_dependencies and not (root / 'lib' / 'msal' / '__init__.py').is_file():
        raise ValueError('Install Windows dependencies into lib before building a release.')
    out = root / 'dist' / ASSET
    out.parent.mkdir(exist_ok=True)
    names = ['main.py', 'ui.py', 'plugin.json', 'requirements.txt', 'README.md',
             'LICENSE', 'PUBLISHING.md', 'outlook_flow', 'Images', 'scripts/install.ps1']
    if require_dependencies:
        names.append('lib')
    with ZipFile(out, 'w', ZIP_DEFLATED) as archive:
        for name in names:
            path = root / name
            if not path.exists():
                raise ValueError(f'Missing package file: {name}')
            files = sorted(path.rglob('*')) if path.is_dir() else [path]
            for file in files:
                if file.is_file() and '__pycache__' not in file.parts and file.suffix != '.pyc' and file.name != '.DS_Store':
                    archive.write(file, file.relative_to(root))
    return out


def repository_source(root=ROOT):
    out = root / 'dist' / 'OutlookWorkspace-GitHub-Source.zip'
    out.parent.mkdir(exist_ok=True)
    names = ['.gitignore', '.github', 'Images', 'outlook_flow', 'scripts', 'tests',
             'submission', 'main.py', 'ui.py', 'plugin.json', 'requirements.txt',
             'README.md', 'PUBLISHING.md', 'LICENSE']
    with ZipFile(out, 'w', ZIP_DEFLATED) as archive:
        for name in names:
            path = root / name
            files = sorted(path.rglob('*')) if path.is_dir() else [path]
            for file in files:
                if file.is_file() and '__pycache__' not in file.parts and file.suffix != '.pyc' and file.name != '.DS_Store':
                    archive.write(file, file.relative_to(root))
    return out


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--release', action='store_true', help='Require and include bundled Windows lib dependencies')
    parser.add_argument('--repository-source', action='store_true', help='Build the GitHub upload bundle instead of a plugin installation ZIP')
    args = parser.parse_args()
    if args.release and args.repository_source:
        parser.error('--release and --repository-source are mutually exclusive')
    print(repository_source() if args.repository_source else package(require_dependencies=args.release))

"""Create and push the prepared public GitHub plugin repository."""
import argparse
import subprocess
import sys
from pathlib import Path
from configure_repository import configure

ROOT = Path(__file__).resolve().parent.parent
FILES = ['.gitignore', '.github', 'Images', 'outlook_flow', 'scripts', 'tests',
         'submission', 'main.py', 'ui.py', 'plugin.json', 'requirements.txt',
         'README.md', 'PUBLISHING.md', 'LICENSE']


def run(*command, check=True):
    return subprocess.run(command, cwd=ROOT, check=check)


def publish(repository):
    # Check GitHub access before modifying local Git metadata.
    run('gh', 'auth', 'status')
    subprocess.run(['gh', 'api', 'user', '--jq', '.login'], cwd=ROOT, check=True)
    probe = subprocess.run(['git', 'remote', 'get-url', 'origin'], cwd=ROOT,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if probe.returncode == 0:
        raise ValueError('An origin remote already exists. Use the manual publishing steps in PUBLISHING.md.')
    configure(repository)
    run(sys.executable, '-m', 'unittest', 'discover', '-s', 'tests', '-v')
    run(sys.executable, 'scripts/package.py')
    if not (ROOT / '.git').exists():
        run('git', 'init', '-b', 'main')
    branch = subprocess.check_output(['git', 'branch', '--show-current'], cwd=ROOT, text=True).strip()
    if branch != 'main':
        raise ValueError('Switch to main before creating the plugin repository.')
    run('git', 'add', '--', *FILES)
    staged = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=ROOT)
    if staged.returncode == 1:
        run('git', 'commit', '-m', 'Add Outlook Workspace Flow Launcher plugin')
    elif staged.returncode != 0:
        raise ValueError('Unable to inspect staged changes.')
    try:
        run('gh', 'repo', 'create', repository, '--public', '--source', '.',
            '--remote', 'origin', '--push', '--description',
            'Outlook email, AI drafts, Calendar and Microsoft To Do for Flow Launcher')
    except subprocess.CalledProcessError:
        print('If GitHub already created the repository, finish manually:', file=sys.stderr)
        print(f'git remote add origin https://github.com/{repository}.git  # only if origin is missing', file=sys.stderr)
        print('git push -u origin main', file=sys.stderr)
        raise
    print(f'Created https://github.com/{repository}')
    print('After CI and Windows acceptance checks pass: git tag v1.0.0 && git push origin v1.0.0')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('repository')
    args = parser.parse_args()
    try:
        publish(args.repository)
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        print(f'Publishing stopped: {exc}', file=sys.stderr)
        sys.exit(1)

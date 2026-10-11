import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from scripts.configure_repository import configure
from scripts.package import ROOT, package, repository_source


class DistributionTests(unittest.TestCase):
    def test_store_manifest_matches_plugin_and_release_asset(self):
        plugin = json.loads((ROOT / 'plugin.json').read_text())
        submission = json.loads(next((ROOT / 'submission').glob('*.json')).read_text())
        for key in ['ID', 'Name', 'Description', 'Author', 'Version', 'Language', 'Website']:
            self.assertEqual(plugin[key], submission[key])
        self.assertEqual(submission['UrlDownload'], plugin['Website'] + '/releases/download/v' + plugin['Version'] + '/Flow.Launcher.Plugin.OutlookWorkspace.zip')
        self.assertTrue(submission['IcoPath'].startswith('https://cdn.jsdelivr.net/gh/'))

    def test_source_package_root_is_installable_and_no_local_data(self):
        with ZipFile(package()) as archive:
            names = archive.namelist()
            self.assertIn('plugin.json', names)
            self.assertIn('main.py', names)
            self.assertIn('Images/app.png', names)
            self.assertFalse(any('.git/' in name or '__pycache__' in name or name.endswith('.env') for name in names))
            plugin = json.loads(archive.read('plugin.json'))
            self.assertIn(plugin['ExecuteFileName'], names)
            self.assertIn(plugin['IcoPath'].replace('\\', '/'), names)

    def test_release_requires_bundled_dependencies(self):
        if not (ROOT / 'lib').exists():
            with self.assertRaisesRegex(ValueError, 'Windows dependencies'):
                package(require_dependencies=True)

    def test_github_upload_contains_workflows_and_submission(self):
        with ZipFile(repository_source()) as archive:
            self.assertIn('.github/workflows/release.yml', archive.namelist())
            self.assertIn('.github/workflows/ci.yml', archive.namelist())
            self.assertTrue(any(name.startswith('submission/') for name in archive.namelist()))
            self.assertFalse(any(name.startswith(('lib/', 'dist/', '.git/')) for name in archive.namelist()))

    def test_configure_repository_keeps_plugin_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'plugin.json').write_text((ROOT / 'plugin.json').read_text())
            target = configure('new-owner/new-plugin', root)
            manifest = json.loads(target.read_text())
            self.assertEqual(manifest['Website'], 'https://github.com/new-owner/new-plugin')
            self.assertEqual(manifest['ID'], json.loads((ROOT / 'plugin.json').read_text())['ID'])

    def test_reject_invalid_repository(self):
        for repository in ['https://github.com/owner/repo', '../bad', 'owner/repo/extra']:
            with self.assertRaises(ValueError):
                configure(repository)


if __name__ == '__main__':
    unittest.main()

# GitHub and Flow Launcher publishing

Prepared repository: `armanislam-07/Flow.Launcher.Plugin.OutlookWorkspace`.

The local repository files are ready. Remote creation requires an authenticated GitHub CLI and internet access. If using another owner/repository, run `python3 scripts/configure_repository.py OWNER/REPOSITORY` first.

## Create the GitHub repository

From this project folder on your own machine:

```sh
gh auth login -h github.com
python3 scripts/publish_repository.py armanislam-07/Flow.Launcher.Plugin.OutlookWorkspace
```

The helper validates tests and metadata, initializes a `main` Git repository if needed, commits only the listed plugin project files, creates a **public** GitHub repository, and pushes `main`. It refuses to run when `origin` already exists, so it cannot silently replace an existing remote. Git needs your author identity configured. If GitHub creation succeeds but a push fails, follow the printed manual commands; rerunning is intentionally blocked once an origin exists.

Alternative: create an empty public repository through GitHub's website, upload the contents of `dist/OutlookWorkspace-GitHub-Source.zip` (extract first), including the `.github` directory, then commit them to `main`. The source ZIP is an upload bundle, not a Flow installation ZIP.

## Publish the installable release

First let **Validate plugin** pass on GitHub, and perform the Windows/account acceptance checks in README.md. Then:

```sh
git tag v1.0.0
git push origin v1.0.0
```

The tag-triggered Windows workflow tests the plugin, installs Python 3.11 Windows dependencies into `lib/`, and publishes `Flow.Launcher.Plugin.OutlookWorkspace.zip` to the `v1.0.0` GitHub Release. `plugin.json` sits at the ZIP root, as Flow expects. The workflow also uploads the generated store manifest as an Actions artifact.

Manual install once that release exists:

```text
pm install https://github.com/armanislam-07/Flow.Launcher.Plugin.OutlookWorkspace/releases/download/v1.0.0/Flow.Launcher.Plugin.OutlookWorkspace.zip
```

Use Flow's Python 3.11 interpreter for the bundled release. It includes native dependencies built for Windows x64 / Python 3.11. Users with a different Python version or architecture should build/install from source with `scripts/install.ps1` instead.

The local `python scripts/package.py` builds a source-only plugin ZIP. `python scripts/package.py --release` requires dependencies already installed into `lib` on Windows and includes them. Do not upload a macOS-built dependency bundle as the Windows release.

## Submit to the Flow plugin store

1. Confirm the public GitHub Release exists and the release asset installs on Windows.
2. Fork [Flow-Launcher/Flow.Launcher.PluginsManifest](https://github.com/Flow-Launcher/Flow.Launcher.PluginsManifest).
3. Copy `submission/OutlookWorkspace-835A8C2B352E4CA9B4FEF5EE3AE05791.json` into that repository's `plugins/` directory.
4. Open a pull request to its default branch, describing the features, the required Microsoft app registration and optional OpenAI API key, and the Windows validation performed.

This project does not automatically submit a store pull request. Store listing depends on Flow's review. For later versions, bump `plugin.json` Version, rerun `configure_repository.py` for the same repository, commit, and push a matching `vX.Y.Z` tag.

Official format and requirements: [Flow plugin submission instructions](https://github.com/Flow-Launcher/Flow.Launcher.PluginsManifest#how-to-submit-your-plugin), [Python build/release guide](https://github.com/Flow-Launcher/docs/blob/main/py-setup-project.md).

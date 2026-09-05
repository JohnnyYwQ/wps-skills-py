# Windows Word Session Setup

The repository contains a production Windows Word Session Host and Adapter for a thirteen-Action surface, plus a standalone Word Application Skill. Document execution runs directly on the Windows machine that has WPS; there is no separate Windows Host service to install.

## Requirements

- Python 3.8 or newer.
- Windows with WPS Writer registered as `KWPS.Application`.
- Native Windows PowerShell under `%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe`. The Runtime does not fall back to the slower WOW64 host.
- No third-party Python package or external service.

The local unit suite itself does not start WPS, COM, or PowerShell.

## Assemble and install the Skill

From the repository root:

```bash
python scripts/build_word_skill.py --output build/skills/wps-word
```

The destination must not already exist. The complete output has this layout:

```text
wps-word/
  SKILL.md
  agents/openai.yaml
  references/
  scripts/word.py
  runtime/
    files.sha256.json
    src/main/python/wps_skills/
    src/main/resources/wps_skills/word/windows/
```

Copy the complete `wps-word` directory to the Skill location supported by the target agent. Do not install only `SKILL.md` or only the source resources directory: the deployed entry point needs the bundled Runtime. To use this Skill on another execution host, place the complete directory on that host too and use that host's paths. No machine names, SSH credentials, or scheduled tasks are embedded.

Verify installation without starting WPS:

```powershell
python "C:\path\to\wps-word\scripts\word.py" --app word --index
python "C:\path\to\wps-word\scripts\word.py" --app word --resolve openDocument inspectDocument writeContent save
```

Successful resolution exits 0; a `partial` or `failed` batch exits 2 while still reporting every requested Action. An unavailable application exits 4 without publishing another application's contracts.

Follow `SKILL.md` and `references/session.md` to execute a task with the Python Session Client. If visible WPS output is needed, execute in the logged-in user's desktop session. SSH execution by itself does not establish desktop visibility; remote desktop launch is environment-specific and is not part of the Skill installer.

## Local verification

From the repository root, run:

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

In Windows PowerShell, set the same source root with:

```powershell
$env:PYTHONPATH = "src/main/python"
python -m unittest discover -s src/test/python -p "test_*.py"
```

## Windows production Session

Run this command inside the repository on the Windows WPS machine:

```bash
python scripts/call.py --session --app word
```

It emits `session.ready` on stdout, then accepts strict JSONL Action Requests. A normal existing-file flow is `openDocument`, any supported required Actions, final `inspectDocument`, explicit `save` when Persistence Intent requires it, then `{"control":"close"}`. The production Set also supports `findContent`, `replaceContent`, `insertTable`, `insertImage`, `setHeaderFooter`, `setPageLayout`, `insertBreak`, and `exportPdf`; only `saveAs` remains deferred. The Runtime keeps one exact document and one owned, console-hidden PowerShell bridge for the full Session. Excel and PPT still fail construction before `session.ready`.

## Admitting more capability

An additional Word target Action, or a future Excel/PPT slice, is admitted only after it owns all of the following:

1. A complete Application Contract Set and generated Action Index.
2. An exact-document Application Adapter and controller at the shared seam.
3. Binding, persistence, handler, and real-WPS verification evidence.
4. Its own independently discoverable Application Skill.

Do not add a placeholder Skill, empty production Contract Set, fake production Adapter, global Action Manifest, or compatibility wrapper around the removed execution model.

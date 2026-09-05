# WPS Automation Foundation

This repository contains the shared WPS Action Session foundation and a production Windows Word slice, including an independently installable `wps-word` Application Skill.

## What exists

- An application-scoped `ActionSession` with immutable one-document binding, closed Controller Result handling, exact-document dispatch, and idempotent cleanup.
- A canonical JSONL `SessionHost` with strict Action Request decoding, lifecycle records, per-Action/session timing journals, traced request rejection, and terminal ordering.
- A complete, validated Word target Contract Set and compact Action Index for fourteen designed Actions.
- A production Word Application Contract Set containing thirteen Actions: document create/open; structured write, inspect, find, and replace; table and image insertion; header/footer, page-layout, and break changes; in-place save; and PDF export. Every advertised required Action has a real handler. Only the deferred `saveAs` target Action remains absent.
- A Session-owned suspended-process launcher with Windows Job Object containment, one lazy native `System32` Windows PowerShell bridge, stable-file-identity acquisition, and a cross-process guard/Lease/quarantine coordinator.
- Real structured Word writing, including separate Western and East Asian run fonts, plus bounded search/replacement, tables, embedded images, headers/footers, layout, page/section breaks, and PDF export, with revision-aware results and operation-specific read-back verification.
- Hidden bridge launch at both process layers: Windows creates the child with `CREATE_NO_WINDOW`, and native PowerShell is also given `-WindowStyle Hidden`, so automation does not open a console window.
- Word establishment makes WPS visible and activates only the exact created or opened document. A newly created WPS application uses normal, not maximized, outer-window state. On either create or attach, a genuinely tiny top-level frame is repaired to a centered 80% of its monitor work area; an already reasonable user window is left alone. Later Actions still dispatch through the retained binding rather than window focus.
- A `scripts/call.py --session --app word` production entry point on Windows. Excel and PPT remain unavailable and fail before `session.ready`.
- Side-effect-free `--app word --index` and `--app word --resolve ACTION...` discovery, generated from the production Contract Set.
- A reusable Python Session Client that preserves terminal Action errors, serializes calls, and separates document results from cleanup outcomes.
- A Word Skill under `src/main/resources/skills/wps-word`, with task guidance, executable client examples, and standalone assembly.
- Local fake-driven conformance tests plus ignored bounded Windows/WPS evidence for the real bridge and production Session Host.

The previous combined Skill, global Manifest and discovery CLI, multi-application Runtime, WPS controllers, Linux/OpenXML backends, and live harnesses were removed by [ADR 0016](docs/adr/0016-cut-over-without-legacy-runtime-compatibility.md). They are recoverable from Git history but are not compatibility interfaces.

## Verify the foundation

Python 3.8 or newer is sufficient:

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

The suite exercises local fakes, real subprocess protocol channels, and relocated Skill distributions. No WPS installation or external account is required.

## Build and use the Word Skill

```bash
python scripts/build_word_skill.py
python build/skills/wps-word/scripts/word.py --app word --index
python build/skills/wps-word/scripts/word.py --app word --resolve createDocument writeContent inspectDocument
```

The build creates `build/skills/wps-word/`, containing `SKILL.md`, references, the thin `scripts/word.py` entry point, and a snapshot of the Python Runtime and PowerShell resources. Copy this complete directory into the target agent's Skill directory. The source tree remains the only maintained implementation; the build includes a SHA-256 file inventory and refuses to overwrite an existing destination. Use `--output <new-directory>/wps-word` for another build.

Read [the Word Skill](src/main/resources/skills/wps-word/SKILL.md) for document intent, discovery, execution, verification, persistence, and failure handling. Its [Session guide](src/main/resources/skills/wps-word/references/session.md) includes a Python task example using `open_session()` and `client.call(address, params)`; the caller decides each next Action after the prior response. Discovery works on macOS/Linux too, while document execution runs on the Windows WPS host.

The source Skill's `scripts/word.py` also works directly from its source location. A deployed Skill uses its bundled Runtime and requires no repository checkout or third-party Python package.

On the Windows WPS host, start a Word Session with:

```bash
python scripts/call.py --session --app word
```

The Host emits `session.ready`, accepts newline-delimited Action Requests, and reuses one exact live Word document and one bridge until `{"control":"close"}`. `saveAs` remains outside the production Contract Set because its gap-free destination Lease migration is still deferred.

Protocol v1 remains closed: timing is diagnostic rather than an extra Action Response field. The `traceLog` in each Action Response records that Action's `elapsedMs`; the Session `traceLog` ends with `sessionElapsedMs`, `actionExecutionElapsedMs`, `cleanupElapsedMs`, and `actionCount` so wall time and actual Action execution are not confused.

Normal Session cleanup deliberately leaves the document open. A controlled debug run that creates disposable content may opt into test-only cleanup:

```bash
python scripts/call.py --session --app word --debug-close-created-document
```

That flag discards and closes only a document created by that Session. It never closes a document acquired through `openDocument`, is not a Word Action, and must not be used when the newly created document contains content that should be retained.

## Project layout

The repository uses Java-style source sets while retaining Python packages:

```text
src/
  main/
    python/wps_skills/
      cli/          # production assembly
      client/       # caller-side Session Protocol and process lifetime
      core/         # application-independent Action Session Core
      host/         # JSONL Session Host
      word/         # Word contracts, handlers, and Adapter
      windows/      # Windows process, coordination, and bridge adapters
    resources/
      wps_skills/word/windows/  # PowerShell/WPS bridge and Action resources
      skills/wps-word/         # Skill source, references, and thin entry point
  test/
    python/tests/   # tests mirror the production modules
    resources/      # non-production capability evidence
scripts/            # thin repository entry points only
docs/               # domain docs, ADRs, and implementation notes
build/              # ignored logs, probes, and test output
```

Tests use a separate `tests` namespace because a second top-level Python package named `wps_skills` would shadow the production package during discovery.

## Design sources

- `CONTEXT.md` defines canonical domain language.
- `docs/adr/` records active architecture decisions.
- `docs/word-action-contracts.md` distinguishes the fourteen-Action Target Contract Portfolio from the thirteen-Action production Application Contract Set.
- `docs/word-adapter-boundary.md` describes the Adapter/Backend, process ownership, and coordination seams.
- `docs/word-action-migration.md` preserves the historical Word capability inventory and future migration evidence.
- `docs/adr/0019-use-java-style-source-sets-around-python-packages.md` records the source-set and Python namespace trade-off.
- `src/test/resources/wps_skills/word/type_library/wps_writer_api.py` is capability evidence only and is never imported by the Runtime.

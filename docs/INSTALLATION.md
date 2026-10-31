# Installation

## 1. Prerequisites

| Need | Why | Check |
|---|---|---|
| Windows 10/11 | SOLIDWORKS and its COM API are Windows-only | |
| SOLIDWORKS 2023, 2024 or 2025 | the scripts drive it through COM (tested on 2025 SP1.2) | Help > About SOLIDWORKS |
| Python 3.9 or newer | runs the build scripts | `python --version` |
| `pywin32` | Python ↔ COM bridge | `pip install pywin32`, then `python -c "import win32com.client"` |
| Claude Code | loads the skill | [code.claude.com](https://code.claude.com) (CLI, desktop app, or VS Code / JetBrains extension) |

Optional: `matplotlib` (outline text labels in `swlib.text_outline`) and `Pillow` (PNG renders).

Check that Python can see your SOLIDWORKS type library (SOLIDWORKS doesn't need to be running):

```bash
python tools/sw_signature.py IModelDoc2.FeatureByPositionReverse
```

The first line shows the type library version: 33 = 2025, 32 = 2024, 31 = 2023.

## 2. Install the skill

### Option A: plugin from this repository (recommended)

```text
/plugin marketplace add Nilove2018/solidworks-api-skill
/plugin install solidworks-api@solidworks-api-skill
```

You can run the same commands from a terminal as `claude plugin marketplace add ...` and
`claude plugin install ...`. The plugin skill is called `/solidworks-api:solidworks-api`.

- **Update:** `/plugin marketplace update solidworks-api-skill`
- **Remove:** `/plugin uninstall solidworks-api@solidworks-api-skill`

From a local clone, before you publish anything:

```bash
claude plugin marketplace add ./solidworks-api-skill
claude plugin install solidworks-api@solidworks-api-skill
```

### Option B: personal skill (one file)

Copy `skills/solidworks-api/SKILL.md` to:

| Scope | Path |
|---|---|
| You, in every project | `~/.claude/skills/solidworks-api/SKILL.md` (Windows: `C:\Users\<you>\.claude\skills\solidworks-api\SKILL.md`) |
| Everyone working in one project | `<project>/.claude/skills/solidworks-api/SKILL.md` (commit it) |

```powershell
# PowerShell, from the folder that contains the clone
New-Item -ItemType Directory -Force "$HOME\.claude\skills\solidworks-api" | Out-Null
Copy-Item .\solidworks-api-skill\skills\solidworks-api\SKILL.md "$HOME\.claude\skills\solidworks-api\SKILL.md"
```

The skill is called `/solidworks-api`. To update it, copy the file again.

### Option C: other AI coding agents

`SKILL.md` is plain Markdown with YAML front matter. Point your agent at it, for example:
- Cursor: copy it into `.cursor/rules/solidworks-api.md`
- Codex / generic agents: reference it from `AGENTS.md` ("Before writing SOLIDWORKS API code,
  read `SKILL.md`")
- GitHub Copilot: include it in `.github/copilot-instructions.md`

## 3. Check that it works

1. Start a **new** Claude Code session. Skills are discovered when a session starts.
2. Ask: *"Which skills do you have for SOLIDWORKS?"* Claude should name `solidworks-api`.
3. Optional end-to-end test, with SOLIDWORKS open and your work saved:

   ```bash
   python examples/01_parametric_plate.py
   ```

   The script should finish with `ALL CHECKS PASSED` and create `examples/out/Parametric_Plate.SLDPRT`.

## 4. Copy the Python layer into your own projects

The skill tells Claude to reuse `swlib.py`. Copy `examples/swlib.py` into your project's
`Scripts/` folder (or keep a path to it), then:

```python
import swlib as L
sw = L.SwApp()                       # attaches to the running SOLIDWORKS (or starts it)
p = L.Part(sw, "Bracket")
p.sketch("z", 0, "SK_BASE"); p.rect(0, 0, 40, 30, r=3); p.extrude("BASE", 5)
p.save(r"C:\work\Bracket.SLDPRT")
```

To force a specific SOLIDWORKS version, set `SW_TLB_MAJOR` (e.g. `set SW_TLB_MAJOR=32` for 2024).

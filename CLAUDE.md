# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

Teaching material for the Quality Management course at HWR Berlin: a full **DMAIC cycle** (Define, Measure, Analyze, Improve, Control) executed around a physical catapult experiment. The repo ships two coupled artifacts:

1. **`DMAIC_Katapult_Versuch.ipynb`** — a Colab-targeted notebook students run.
2. **`catapult/`** — `statapult`, a standalone CLI/Python simulator used to generate plausible experiment data (for testing, demos, or when the physical catapult is unavailable).

These two pieces live in the same repo but are otherwise independent: the notebook does not import `statapult`, and `statapult` does not import `helper.py`.

## Critical Workflow: The Notebook is Generated, Not Hand-Edited

`DMAIC_Katapult_Versuch.ipynb` is a **build artifact**. Never edit it directly — changes will be overwritten.

- Cell contents live as string constants in `scripts/notebook_builder/{intro,define,measure,analyze,improve,control,closing}.py`.
- `scripts/notebook_builder/cells.py` defines helpers (`md`, `code`, `colab_code`, `factor_def_cell`, `phase_export_cell`) and notebook-level metadata.
- `scripts/build_notebook.py` concatenates `PHASE_MODULES` via `build_notebook()` and writes the `.ipynb`.

### Drei-Schritt-Loop für jede Änderung

1. **Quelle anpassen** — `helper.py` oder `scripts/notebook_builder/*.py`.
2. **Testen** — Logik isoliert prüfen (Python-REPL gegen `helper.py`, `pytest catapult/tests/` falls `statapult` betroffen), `python scripts/build_notebook.py --check` gegen den letzten Commit.
3. **Notebook regenerieren** — `python scripts/build_notebook.py`.

Das `.ipynb` wird niemals direkt editiert. Zusätzlich: `helper.py` wird in Colab zur Laufzeit von GitHub geladen — Änderungen wirken für Studierende erst **nach Push auf `main`**.

### Phasen-Verantwortlichkeiten

- **Faktoren werden in DEFINE definiert** (Felder `name, einheit, low, high, centerpoint_moeglich`) und in `projekt.faktoren` abgelegt. ANALYZE verfeinert diese Liste nur noch (Teilmenge auswählen, Centerpoints-Flag anpassen) und schreibt das Ergebnis in `projekt.faktoren_doe`. Nachfolgender DoE-Code liest `faktoren_doe` mit Fallback auf `faktoren`.
- **Zielweite wird in DEFINE vom Nutzer eingetragen** (Default 300.0 cm). Keine Hash-/Random-Zuweisung mehr. Die Min/Max-Vermessung in DEFINE liefert den realistischen Bereich, innerhalb dessen die Zielweite sinnvoll liegt.
- **USL/LSL** (Spec Limits, in CONTROL für Cpk) bleiben `zielweite ± toleranz`.

### Commands

```bash
# Regenerate the notebook from Python sources (run after editing any phase module)
python scripts/build_notebook.py

# Verify the committed .ipynb matches the generator (CI-style check)
python scripts/build_notebook.py --check

# Show unified diff if drift is detected
python scripts/build_notebook.py --diff

# Write to an alternate path
python scripts/build_notebook.py --out /tmp/nb.ipynb
```

`--check` normalizes volatile Colab fields (cell `id`, `metadata.colab.provenance`) and collapses `source` lists to strings before comparing, so post-session Colab noise doesn't cause false drift.

Requires `nbformat` (install via `pip install nbformat`).

## `helper.py` — The Notebook's Statistics Backbone

Single ~3,300-line module that contains **all** statistical logic, plotting, export, and Colab/Drive integration for the notebook. Students never see this code; they interact through form-based cells that call into it.

- Loaded at notebook startup from GitHub (intro cell 3), not installed locally. Editing `helper.py` → commit/push so Colab sessions pick it up.
- Sections are block-delimited (DEFINE / MEASURE / ANALYZE / IMPROVE / CONTROL / EXPORT). The module docstring maps them.
- The central `Projekt` dataclass holds phase state; `exportiere_phase_auf_drive(projekt, phase)` persists plots + `fortschritt.json` to `MyDrive/DMAIC_Katapult/<Gruppe>/<PHASE>/`. Regression models are not pickled — they are rebuilt from saved raw data on load.
- `statsmodels` is imported lazily inside functions that need it, so failures localize to a specific call site.

When modifying `helper.py`, keep the public function signatures stable — the notebook-generator cells call them by name.

## `catapult/` — The Statapult Simulator (separate Python package)

Installable package (`statapult`) with its own `pyproject.toml`, `src/` layout, and 115-test pytest suite. Structure under `src/statapult/`:

- `physics.py` — additive model `d = D_BASE + Σ(aᵢ·xᵢ) + Σ(qᵢ·xᵢ²) + Σ(bᵢⱼ·xᵢ·xⱼ) + noise`
- `noise.py` — layered noise (measurement, setup, rubber-band, release, wind)
- `factors.py` — `STANDARD_FACTORS`, coded↔natural conversion, `to_helper_format()` bridge
- `simulator.py` — `Statapult` orchestrator (`.shoot()`, `.batch()`)
- `excel_filler.py` — auto-detects and fills the notebook's Excel templates (MSA / DoE / Konfirmation)
- `cli.py` — subcommands: `shoot`, `batch`, `msa`, `fill`, `control`, `info`
- `config.py` + `defaults.yaml` — tunable physics/noise parameters

### Commands

```bash
# Install (editable) from repo root
pip install -e ./catapult                # runtime only
pip install -e "./catapult[batch]"       # + pandas/openpyxl for batch + Excel
pip install -e "./catapult[dev]"         # + pytest

# Run the suite
cd catapult && python -m pytest tests/ -v

# Run a single test module / test
python -m pytest catapult/tests/test_physics.py -v
python -m pytest catapult/tests/test_physics.py::test_name -v
```

## Cross-Cutting Notes

- Language: user-facing strings (notebook cells, CLI, README) are in **German**. Keep that convention when editing those surfaces; internal code comments may be English.
- There is no top-level `pyproject.toml` or lockfile. The notebook's runtime deps (`numpy pandas scipy statsmodels matplotlib openpyxl`) are installed ad-hoc in Colab's intro cell; `catapult/` manages its own deps.
- `.gitignore` at repo root is minimal (1 line) — check before adding generated artifacts to the tree.

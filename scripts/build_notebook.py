"""Generate DMAIC_Katapult_Versuch.ipynb from the phase modules.

Usage:
    python scripts/build_notebook.py            # write notebook
    python scripts/build_notebook.py --check    # exit 1 if drift vs. committed file
    python scripts/build_notebook.py --diff     # print unified diff
    python scripts/build_notebook.py --out PATH # write elsewhere
"""
from __future__ import annotations

import argparse
import copy
import difflib
import io
import json
import sys
from pathlib import Path

import nbformat

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUT = REPO_ROOT / "DMAIC_Katapult_Versuch.ipynb"

sys.path.insert(0, str(Path(__file__).resolve().parent))

from notebook_builder import intro, define, measure, analyze, improve, control, closing  # noqa: E402
from notebook_builder.cells import build_notebook  # noqa: E402


PHASE_MODULES = [intro, define, measure, analyze, improve, control, closing]


def assemble():
    all_cells = []
    for mod in PHASE_MODULES:
        all_cells.extend(mod.cells())
    return build_notebook(all_cells)


def serialize(nb) -> str:
    buf = io.StringIO()
    nbformat.write(nb, buf)
    text = buf.getvalue()
    if not text.endswith("\n"):
        text += "\n"
    return text


def normalize_for_check(nb_json: dict) -> dict:
    """Strip volatile Colab fields and collapse source to a single string so
    --check ignores post-session drift and list-vs-string source variations."""
    nb = copy.deepcopy(nb_json)
    colab = nb.get("metadata", {}).get("colab", {})
    colab.pop("provenance", None)
    for cell in nb.get("cells", []):
        cell.get("metadata", {}).pop("id", None)
        cell.pop("id", None)
        src = cell.get("source")
        if isinstance(src, list):
            cell["source"] = "".join(src)
    return nb


def check(out_path: Path, show_diff: bool) -> int:
    generated = assemble()
    gen_text = serialize(generated)
    gen_norm = normalize_for_check(json.loads(gen_text))

    if not out_path.exists():
        print(f"❌ {out_path} does not exist; run without --check first")
        return 1

    committed = json.loads(out_path.read_text(encoding="utf-8"))
    com_norm = normalize_for_check(committed)

    gen_s = json.dumps(gen_norm, indent=1, ensure_ascii=False, sort_keys=True)
    com_s = json.dumps(com_norm, indent=1, ensure_ascii=False, sort_keys=True)

    if gen_s == com_s:
        print(f"✅ {out_path.name} matches generator output")
        return 0

    print(f"❌ {out_path.name} differs from generator output")
    if show_diff:
        diff = difflib.unified_diff(
            com_s.splitlines(keepends=True),
            gen_s.splitlines(keepends=True),
            fromfile=f"{out_path.name} (committed)",
            tofile=f"{out_path.name} (generated)",
            n=3,
        )
        sys.stdout.writelines(diff)
    else:
        print("   (run with --diff for details)")
    return 1


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    p.add_argument("--check", action="store_true", help="exit 1 on drift, do not write")
    p.add_argument("--diff", action="store_true", help="like --check, print unified diff")
    p.add_argument("--out", type=Path, default=DEFAULT_OUT, help="output path")
    args = p.parse_args()

    if args.check or args.diff:
        return check(args.out, show_diff=args.diff)

    nb = assemble()
    with open(args.out, "w", encoding="utf-8") as f:
        nbformat.write(nb, f)
    print(f"wrote {args.out} ({len(nb['cells'])} cells)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

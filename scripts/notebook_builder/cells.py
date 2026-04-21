"""Cell constructors and notebook-level metadata for DMAIC notebook build."""
from __future__ import annotations

import nbformat
from nbformat.notebooknode import NotebookNode


NOTEBOOK_METADATA: dict = {
    "colab": {"provenance": [], "toc_visible": True},
    "kernelspec": {"name": "python3", "display_name": "Python 3"},
    "language_info": {"name": "python"},
}


def md(source: str) -> NotebookNode:
    cell = nbformat.v4.new_markdown_cell(source)
    cell["metadata"] = {}
    cell.pop("id", None)
    return cell


def code(source: str, *, cell_view_form: bool = True) -> NotebookNode:
    cell = nbformat.v4.new_code_cell(source)
    cell["metadata"] = {"cellView": "form"} if cell_view_form else {}
    cell["execution_count"] = None
    cell["outputs"] = []
    cell.pop("id", None)
    return cell


def colab_code(title: str, body: str) -> NotebookNode:
    return code(f"#@title {title}\n{body}")


def phase_export_cell(phase: str) -> NotebookNode:
    title = f"💾 Ergebnisse bis {phase} auf Google Drive speichern"
    body = f'helper.exportiere_phase_auf_drive(projekt, "{phase}")\n'
    return colab_code(title, body)


def factor_def_cell(
    n: int,
    *,
    name: str = "",
    unit: str = "",
    low: float = 0.0,
    high: float = 0.0,
    optional: bool = False,
    optional_hint: str = "",
    no_factor_msg: str = "",
) -> NotebookNode:
    """Generate one 'Faktor N definieren' code cell.

    For n in (1, 2, 3): non-optional variant.
    For n in (4, 5): optional variant with empty-name guard.
    """
    if optional:
        title = f"📝 Faktor {n} ({optional_hint})"
    else:
        title = f"📝 Faktor {n} definieren"

    name_lit = f'"{name}"'
    unit_lit = f'"{unit}"'
    p = f"f{n}"
    head = (
        f'{p}_name = {name_lit} #@param {{type:"string"}}\n'
        f'{p}_einheit = {unit_lit} #@param {{type:"string"}}\n'
        f'{p}_low = {low} #@param {{type:"number"}}\n'
        f'{p}_high = {high} #@param {{type:"number"}}\n'
        f'{p}_centerpoint = True #@param {{type:"boolean"}}\n'
        f'\n'
    )
    faktor_expr = (
        f'{{"name": {p}_name, "einheit": {p}_einheit, "low": {p}_low, '
        f'"high": {p}_high, "centerpoint_moeglich": {p}_centerpoint}}'
    )
    cp_expr = (
        f'"stetig → Centerpoint möglich" if {p}_centerpoint '
        f'else "zweistufig → kein Centerpoint"'
    )
    ok_print = (
        f'    print(f"✅ Faktor {n}: {{{p}_name}} '
        f'[{{{p}_low}} – {{{p}_high}} {{{p}_einheit}}] ({{_cp}})")'
    )

    if optional:
        body = (
            head
            + f'if {p}_name.strip():\n'
            + f'    faktor{n} = {faktor_expr}\n'
            + f'    _cp = {cp_expr}\n'
            + ok_print + "\n"
            + f'else:\n'
            + f'    faktor{n} = None\n'
            + f'    print("ℹ️ {no_factor_msg}")\n'
        )
    else:
        body = (
            head
            + f'faktor{n} = {faktor_expr}\n'
            + f'_cp = {cp_expr}\n'
            + f'print(f"✅ Faktor {n}: {{{p}_name}} '
            f'[{{{p}_low}} – {{{p}_high}} {{{p}_einheit}}] ({{_cp}})")\n'
        )

    return colab_code(title, body)


def build_notebook(cells: list[NotebookNode]) -> NotebookNode:
    nb = nbformat.v4.new_notebook()
    nb["metadata"] = dict(NOTEBOOK_METADATA)
    nb["nbformat"] = 4
    nb["nbformat_minor"] = 0
    nb["cells"] = list(cells)
    return nb

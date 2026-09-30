"""Workbook inventory, dependency graph, descriptive rules and migration order."""

from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter

from .formula import CELL, VOLATILE, pattern, tokenize

BANNER = "SYNTHETIC FIXTURE ANALYSIS ONLY. Read-only and descriptive; no financial audit or correctness certification. Migration ordering is a proposal, not a guarantee."
SPAGHETTI_THRESHOLD = 3
RANGE_LIMIT = 100000


def cell_id(sheet, coordinate):
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'!{coordinate}"


def parse_workbook(path):
    """Read formulas without writing, evaluating or trusting cached results."""
    workbook = load_workbook(path, read_only=False, data_only=False, keep_links=True)
    cells, names, issues = {}, [], []
    try:
        for sheet in workbook:
            for row in sheet:
                for cell in row:
                    if cell.value is None:
                        continue
                    value = cell.value
                    formula = value if cell.data_type == "f" and isinstance(value, str) else None
                    if not isinstance(value, (str, int, float, bool, date, datetime)):
                        issues.append({"cell": cell_id(sheet.title, cell.coordinate), "reason": "Unsupported cell value or array/data-table formula"})
                        value = type(value).__name__
                    if isinstance(value, (date, datetime)):
                        value = value.isoformat()
                    cells[cell_id(sheet.title, cell.coordinate)] = {
                        "sheet": sheet.title, "coordinate": cell.coordinate,
                        "value": value, "formula": formula,
                        "analysis": tokenize(formula) if formula else None}
            for name in sheet.defined_names.values():
                names.append({"name": name.name, "scope": sheet.title, "definition": name.attr_text})
            if sheet.tables:
                issues.append({"cell": cell_id(sheet.title, "A1"), "reason": "Structured tables require manual review; cell values remain inventoried"})
        for name in workbook.defined_names.values():
            names.append({"name": name.name, "scope": None, "definition": name.attr_text})
        if workbook._external_links:
            issues.append({"cell": None, "reason": "External workbook links cannot be resolved"})
        label = cells.get(cell_id("README", "A1"), {}).get("value", "")
        return {"file": Path(path).name, "synthetic_labelled": str(label).startswith("SYNTHETIC"),
                "sheets": [sheet.title for sheet in workbook], "cells": dict(sorted(cells.items())),
                "names": sorted(names, key=lambda n: (n["name"], n["scope"] or "")),
                "issues": issues}
    finally:
        workbook.close()


def resolve(reference, sheet, inventory, stack=()):
    """Resolve finite A1 rectangles and workbook/sheet scoped names."""
    names = inventory["names"]
    scope = sheet
    address = reference
    if "!" in reference:
        prefix, address = reference.rsplit("!", 1)
        scope = prefix.strip("'").replace("''", "'")
        if scope not in inventory["sheets"]:
            raise ValueError(f"Unsupported or missing sheet: {prefix}")
    candidates = [n for n in names if n["name"].lower() == address.lower() and n["scope"] in (None, scope)]
    if candidates:
        name = sorted(candidates, key=lambda n: n["scope"] is None)[0]
        key = (name["scope"], name["name"])
        if key in stack:
            raise ValueError(f"Recursive defined name: {address}")
        definition = name["definition"].lstrip("=")
        return resolve(definition, name["scope"] or scope, inventory, stack + (key,))
    endpoints = address.split(":")
    if len(endpoints) > 2 or not all(CELL.fullmatch(part) for part in endpoints):
        raise ValueError(f"Unsupported reference or defined name: {reference}")
    min_col, min_row, max_col, max_row = range_boundaries(address)
    if max_col > 16384 or max_row > 1048576 or min_col > max_col or min_row > max_row:
        raise ValueError(f"Invalid Excel range: {reference}")
    if (max_col - min_col + 1) * (max_row - min_row + 1) > RANGE_LIMIT:
        raise ValueError(f"Range exceeds declared expansion limit {RANGE_LIMIT}: {reference}")
    return [cell_id(scope, f"{get_column_letter(col)}{row}")
            for row in range(min_row, max_row + 1) for col in range(min_col, max_col + 1)]


def components(nodes, adjacency):
    """Iterative Kosaraju avoids recursion limits on long dependency chains."""
    seen, order = set(), []
    for start in sorted(nodes):
        if start in seen:
            continue
        stack = [(start, False)]
        while stack:
            node, finished = stack.pop()
            if finished:
                order.append(node)
                continue
            if node in seen:
                continue
            seen.add(node)
            stack.append((node, True))
            stack.extend((n, False) for n in reversed(sorted(adjacency[node])) if n not in seen)
    reverse = defaultdict(set)
    for source in nodes:
        for target in adjacency[source]:
            reverse[target].add(source)
    seen, result = set(), []
    for start in reversed(order):
        if start in seen:
            continue
        group, stack = [], [start]
        seen.add(start)
        while stack:
            node = stack.pop()
            group.append(node)
            for other in sorted(reverse[node]):
                if other not in seen:
                    seen.add(other)
                    stack.append(other)
        result.append(sorted(group))
    return sorted(result)


def cycle_path(group, adjacency):
    allowed = set(group)
    start = group[0]
    stack = [(start, [start])]
    seen = set()
    while stack:
        node, path = stack.pop()
        if node in seen:
            continue
        seen.add(node)
        for other in reversed(sorted(adjacency[node] & allowed)):
            if other == start:
                return path + [start]
            if other not in seen:
                stack.append((other, path + [other]))
    raise RuntimeError("Cyclic component has no cycle")


def analyze(path):
    inventory = parse_workbook(path)
    cells = inventory["cells"]
    issues = list(inventory["issues"])
    if not inventory["synthetic_labelled"]:
        issues.append({"cell": None, "reason": "Input is not a labelled synthetic fixture. This tool does not audit real finances."})
    edges, nodes = [], set(cells)
    for identifier, cell in cells.items():
        if not cell["formula"]:
            continue
        for reason in cell["analysis"]["issues"]:
            issues.append({"cell": identifier, "reason": reason})
        dependencies = set()
        for reference in cell["analysis"]["references"]:
            try:
                dependencies.update(resolve(reference, cell["sheet"], inventory))
            except ValueError as exc:
                issues.append({"cell": identifier, "reason": str(exc)})
        for dependency in sorted(dependencies):
            nodes.add(dependency)
            edges.append({"source": dependency, "target": identifier, "formula": cell["formula"]})
    adjacency = defaultdict(set)
    for edge in edges:
        adjacency[edge["source"]].add(edge["target"])
    groups = components(nodes, adjacency)
    cycles = [cycle_path(g, adjacency) for g in groups if len(g) > 1 or g[0] in adjacency[g[0]]]
    graph = {"direction": "precedent -> dependent", "nodes": sorted(nodes),
             "edges": sorted(edges, key=lambda e: (e["source"], e["target"])), "cycles": cycles}
    rules = []

    def rule(kind, citations, description, **details):
        rules.append({"id": f"R{len(rules) + 1:04}", "kind": kind, "cells": sorted(set(citations)),
                      "description": description, **details})

    outputs = set()
    for name in inventory["names"]:
        try:
            members = resolve(name["name"], name["scope"] or inventory["sheets"][0], inventory)
            feeds = sorted({target for member in members for target in adjacency[member]})
            rule("named_range", members + feeds, f"Named range {name['name']} defines {name['definition']} and feeds {len(feeds)} formula cells.",
                 name=name["name"], definition=name["definition"], members=members, feeds=feeds)
            if name["name"].lower().startswith("output"):
                outputs.update(members)
        except ValueError as exc:
            issues.append({"cell": None, "reason": f"Defined name {name['name']}: {exc}"})
    for identifier, cell in cells.items():
        if not cell["formula"]:
            continue
        functions = cell["analysis"]["functions"]
        rule("formula", [identifier], f"{identifier} computes {cell['formula']}", formula=cell["formula"])
        for function in functions:
            if function in {"VLOOKUP", "HLOOKUP", "XLOOKUP"}:
                refs = cell["analysis"]["references"]
                resolved = set()
                for ref in refs:
                    try:
                        resolved.update(resolve(ref, cell["sheet"], inventory))
                    except ValueError:
                        pass
                shape = {"VLOOKUP": "searches the first table column and returns a selected column",
                         "HLOOKUP": "searches the first table row and returns a selected row",
                         "XLOOKUP": "searches a lookup array and returns from a corresponding return array"}[function]
                rule("lookup", [identifier] + sorted(resolved), f"{function} {shape}: {cell['formula']}. Match mode and fallback follow the formula.",
                     function=function, references=refs, formula=cell["formula"])
            if function in {"IF", "IFS", "IFERROR"}:
                rule("conditional", [identifier], f"{function} branches according to the conditions and result expressions in {cell['formula']}.", formula=cell["formula"])
            if function in {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"}:
                upstream = sorted(e["source"] for e in edges if e["target"] == identifier)
                rule("aggregation", [identifier] + upstream, f"{function} aggregates precedents into {identifier}: {cell['formula']}", precedents=upstream)
    smells = []

    def smell(kind, citations, explanation, **details):
        smells.append({"id": f"S{len(smells) + 1:04}", "kind": kind, "cells": sorted(set(citations)), "explanation": explanation, **details})

    for cycle in cycles:
        smell("circular_reference", cycle, "Circular dependencies prevent a dependency-safe implementation order; resolve the feedback explicitly.", path=cycle)
    dead = [key for key, c in cells.items() if c["sheet"] != "README" and (c["formula"] or isinstance(c["value"], (int, float))) and not adjacency[key] and key not in outputs]
    if dead:
        smell("dead_cells", dead, "These numeric or formula cells have no dependents and are not declared Output names. Confirm whether they are unused or intended outputs.")
    patterns = defaultdict(list)
    for identifier, cell in cells.items():
        if not cell["formula"]:
            continue
        patterns[pattern(cell["formula"], cell["coordinate"])].append(identifier)
        functions = sorted(set(cell["analysis"]["functions"]) & VOLATILE)
        if functions:
            smell("volatile_functions", [identifier], "Volatile functions can change or recalculate independently of explicit dependencies.", functions=functions)
        if cell["analysis"]["constants"]:
            smell("hardcoded_constants", [identifier], "Embedded literals hide assumptions and should become documented parameters.", constants=cell["analysis"]["constants"])
        sheets = {cells[e["source"]]["sheet"] if e["source"] in cells else e["source"].rsplit("!", 1)[0] for e in edges if e["target"] == identifier}
        sheets.discard(cell["sheet"])
        if len(sheets) >= SPAGHETTI_THRESHOLD:
            smell("cross_sheet_spaghetti", [identifier], f"Formula draws from at least {SPAGHETTI_THRESHOLD} other sheets, coupling separate model areas.", sheets=sorted(sheets), threshold=SPAGHETTI_THRESHOLD)
    for signature, citations in sorted(patterns.items()):
        if len(citations) > 1:
            smell("duplicated_formula_pattern", citations, "Repeated relative formula structure suggests a shared implementation rule.", pattern=signature)
    for issue in issues:
        smell("manual_review", [issue["cell"]] if issue["cell"] else [], issue["reason"])
    index = {node: i for i, group in enumerate(groups) for node in group}
    predecessors = defaultdict(set)
    for edge in edges:
        a, b = index[edge["source"]], index[edge["target"]]
        if a != b:
            predecessors[b].add(a)
    pending, completed, blocked, phases = set(range(len(groups))), set(), set(), []
    while pending:
        ready = sorted(i for i in pending if predecessors[i] <= completed)
        if not ready:
            raise RuntimeError("Component graph must be acyclic")
        phase_cells = sorted(node for i in ready for node in groups[i])
        cycle_groups = [i for i in ready if len(groups[i]) > 1 or groups[i][0] in adjacency[groups[i][0]]]
        blocked.update(i for i in ready if i in cycle_groups or predecessors[i] & blocked)
        phases.append({"phase": len(phases) + 1, "cells": phase_cells,
                       "rules": [r["id"] for r in rules if set(r["cells"]) & set(phase_cells)],
                       "smells_to_address": [s["id"] for s in smells if set(s["cells"]) & set(phase_cells)],
                       "blocked_cells": sorted(node for i in ready if i in blocked for node in groups[i]),
                       "action": "Resolve feedback/manual review and consolidate rules; reimplement only after dependencies are available." if any(i in blocked for i in ready) else "Implement inputs or rules after prior phases; review cited smells."})
        completed.update(ready)
        pending.difference_update(ready)
    plan = {"disclaimer": "Proposed ordering only. Cyclic components and their downstream dependents are blocked until feedback is resolved. Smell remediation is proposed, not performed.",
            "phases": phases, "manual_review_required": bool(issues),
            "unassigned_smells": [s["id"] for s in smells if not s["cells"]]}
    return {"banner": BANNER, "summary": {"file": inventory["file"], "synthetic_labelled": inventory["synthetic_labelled"],
            "sheets": inventory["sheets"], "cell_count": len(cells), "formula_count": sum(bool(c["formula"]) for c in cells.values()),
            "spaghetti_threshold": SPAGHETTI_THRESHOLD}, "inventory": inventory, "graph": graph, "rules": rules, "smells": smells, "plan": plan}

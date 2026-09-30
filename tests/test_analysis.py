import hashlib
import json
import os
import subprocess
import sys

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from sre.analysis import BANNER, analyze, cell_id, parse_workbook, resolve
from sre.cli import markdown, serialize
from sre.fixtures import generate
from sre.formula import pattern, tokenize


def p(address):
    return cell_id("Pricing", address)


@pytest.mark.parametrize("model,cells,formulas", [("clean", 8, 1), ("pricing", 27, 11)])
def test_inventory_and_read_only(models, model, cells, formulas):
    path = models[model]
    before = path.read_bytes()
    result = analyze(path)
    assert result["summary"]["cell_count"] == cells
    assert result["summary"]["formula_count"] == formulas
    assert result["summary"]["synthetic_labelled"]
    assert path.read_bytes() == before
    assert hashlib.sha256(path.read_bytes()).digest() == hashlib.sha256(before).digest()
    if model == "pricing":
        assert result["inventory"]["cells"][p("C2")]["formula"] == "=B2*1.2"
        assert result["inventory"]["cells"][p("J2")]["value"] == 999
        assert {key for key, cell in result["inventory"]["cells"].items() if cell["formula"]} == {p(c) for c in ("C2", "C3", "D2", "E2", "F2", "G2", "H2", "C8", "C9", "C10", "K2")}


def test_full_hand_computed_graph(models):
    graph = analyze(models["pricing"])["graph"]
    expected = set()
    for target, sources in {
        "C2": ["B2"], "C3": ["B3"], "D2": ["B2", "C2"], "E2": ["A2"],
        "F2": ["C2", "C3"], "H2": ["B2"], "C8": ["C9"], "C9": ["C8"],
        "C10": ["C8", "B2"],
    }.items():
        expected.update((p(source), p(target)) for source in sources)
    expected.update((cell_id("Rates", source), p("E2")) for source in ("A2", "A3", "B2", "B3"))
    expected.add((cell_id("Tax", "A2"), p("H2")))
    expected.update((cell_id(sheet, addr), p("K2")) for sheet, addr in [("Rates", "B2"), ("Tax", "A2"), ("Shipping", "A2")])
    assert {(e["source"], e["target"]) for e in graph["edges"]} == expected
    inventory = parse_workbook(models["pricing"])
    assert all(e["formula"] == inventory["cells"][e["target"]]["formula"] for e in graph["edges"])
    assert graph["cycles"] == [[p("C8"), p("C9"), p("C8")]]


def test_clean_graph_and_no_smells(models):
    result = analyze(models["clean"])
    assert result["smells"] == []
    assert result["graph"]["cycles"] == []
    assert {(e["source"], e["target"]) for e in result["graph"]["edges"]} == {
        (cell_id("Inputs", "A2"), cell_id("Calc", "A2")),
        (cell_id("Inputs", "B2"), cell_id("Calc", "A2")),
    }


def test_rules_and_citations(models):
    rules = analyze(models["pricing"])["rules"]
    tax = next(r for r in rules if r.get("name") == "TaxRate")
    assert tax["members"] == [cell_id("Tax", "A2")]
    assert tax["feeds"] == [p("H2"), p("K2")]
    lookup = next(r for r in rules if r["kind"] == "lookup")
    assert lookup["function"] == "VLOOKUP"
    assert lookup["arguments"] == ["A2", "Rates!A2:B3", "2", "FALSE"]
    assert lookup["cells"] == sorted([p("A2"), p("E2")] + [cell_id("Rates", c) for c in ("A2", "A3", "B2", "B3")])
    conditional = next(r for r in rules if r["kind"] == "conditional")
    assert conditional["cells"] == [p("D2")]
    assert conditional["arguments"] == ["B2>15", "C2", "B2"]
    aggregate = next(r for r in rules if r["kind"] == "aggregation")
    assert aggregate["cells"] == [p("C2"), p("C3"), p("F2")]


@pytest.mark.parametrize("kind,expected", [
    ("circular_reference", ["C8", "C9"]),
    ("dead_cells", ["C10", "J2"]),
    ("duplicated_formula_pattern", ["C2", "C3"]),
    ("volatile_functions", ["G2"]),
    ("cross_sheet_spaghetti", ["K2"]),
    ("hardcoded_constants", ["C2", "C3", "C8", "C9", "D2", "E2"]),
])
def test_planted_smells(models, kind, expected):
    findings = [s for s in analyze(models["pricing"])["smells"] if s["kind"] == kind]
    assert sorted({c for s in findings for c in s["cells"]}) == sorted(p(c) for c in expected)
    assert all(s["explanation"] for s in findings)
    if kind == "circular_reference":
        assert findings[0]["path"] == [p("C8"), p("C9"), p("C8")]
    if kind == "hardcoded_constants":
        assert next(s for s in findings if s["cells"] == [p("C2")])["constants"] == ["1.2"]


@pytest.mark.parametrize("model", ["clean", "pricing"])
def test_plan_topology_and_rule_coverage(models, model):
    result = analyze(models[model])
    phases = result["plan"]["phases"]
    assigned = {cell: phase["phase"] for phase in phases for cell in phase["cells"]}
    assert sum(len(phase["cells"]) for phase in phases) == len(assigned)
    assert set(assigned) == set(result["graph"]["nodes"])
    cycle_cells = {cell for cycle in result["graph"]["cycles"] for cell in cycle}
    for edge in result["graph"]["edges"]:
        if edge["source"] in cycle_cells and edge["target"] in cycle_cells:
            assert assigned[edge["source"]] == assigned[edge["target"]]
        else:
            assert assigned[edge["source"]] < assigned[edge["target"]]
    assert {rule for phase in phases for rule in phase["rules"]} == {r["id"] for r in result["rules"]}
    assert {smell for phase in phases for smell in phase["smells_to_address"]} == {s["id"] for s in result["smells"]}
    blocked = {c for phase in phases for c in phase["blocked_cells"]}
    assert blocked == ({p("C8"), p("C9"), p("C10")} if model == "pricing" else set())


def test_fixture_byte_determinism(tmp_path):
    first, second = tmp_path / "one", tmp_path / "two"
    generate(first)
    generate(second)
    for name in ("clean.xlsx", "pricing.xlsx"):
        assert (first / name).read_bytes() == (second / name).read_bytes()


@pytest.mark.parametrize("model", ["clean", "pricing"])
def test_report_determinism_and_banner(models, model):
    a, b = analyze(models[model]), analyze(models[model])
    assert serialize(a).encode() == serialize(b).encode()
    assert markdown(a).encode() == markdown(b).encode()
    assert BANNER in serialize(a) and BANNER in markdown(a)
    for key in ("graph", "rules", "smells", "plan"):
        assert serialize(a[key]) == serialize(b[key])


@pytest.mark.parametrize("formula", [
    "=A1+$B$2", "=-A1+2^3%", '=IF(A1>2,SUM(B1:B2),"none")',
    "=ROUND((A1+B1)/C1,2)", "=TODAY()", "=A1<>FALSE", "=A1&\"x\"", "=1E-3",
])
def test_supported_grammar(formula):
    assert tokenize(formula)["issues"] == []


@pytest.mark.parametrize("formula", ["=A1 B1", "=A1+", "={1,2}", "=SUM(A1,,B1)", "=FOOBAR(A1)", "=()", "=INDIRECT(\"A1\")", "=OFFSET(A1,1,1)"])
def test_unsupported_grammar_flagged(formula):
    assert tokenize(formula)["issues"]


def test_relative_patterns():
    assert pattern("=B2*1.2", "C2") == pattern("=B3*1.2", "C3")
    assert pattern("=$B$2+A2", "C2") == pattern("=$B$2+A3", "C3")
    assert pattern("=B2", "C2") != pattern("=$B2", "C2")
    assert pattern("='A1'!B2", "C2") != pattern("='A2'!B2", "C2")


def test_manual_review_locations_blank_nodes_and_names(tmp_path):
    path = tmp_path / "unsupported.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "O'Brien"
    ws["A1"] = "='O''Brien'!B1"
    ws["A2"] = "=Missing!A1"
    ws["A3"] = "=SUM(A:A)"
    ws["A4"] = '=Table1[Amount]'
    ws["A5"] = '=INDIRECT("B1")'
    ws["A6"] = "=SUM('O''Brien'!A1:XFD1048576)"
    ws["A7"] = "=Loop"
    ws["A8"] = ArrayFormula(ref="A8:A9", text="=A1:A2")
    wb.defined_names.add(DefinedName("Loop", attr_text="Loop"))
    wb.save(path)
    result = analyze(path)
    assert cell_id("O'Brien", "B1") in result["graph"]["nodes"]
    assert result["plan"]["manual_review_required"]
    assert result["graph"]["complete"] is False
    assert result["inventory"]["issues"]
    manual = {c for s in result["smells"] if s["kind"] == "manual_review" for c in s["cells"]}
    assert manual == {cell_id("O'Brien", f"A{i}") for i in range(2, 9)}
    assert result["summary"]["synthetic_labelled"] is False
    assert any("not a labelled synthetic fixture" in s["explanation"] for s in result["smells"])


def test_scoped_and_case_insensitive_references(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"], ws["B1"], ws["C1"] = 2, 3, "=Rate+inputs!A1"
    wb.defined_names.add(DefinedName("Rate", attr_text="'Inputs'!A1"))
    ws.defined_names.add(DefinedName("Rate", attr_text="'Inputs'!B1", localSheetId=0))
    path = tmp_path / "scope.xlsx"
    wb.save(path)
    result = analyze(path)
    assert {(e["source"], e["target"]) for e in result["graph"]["edges"]} == {
        (cell_id("Inputs", "A1"), cell_id("Inputs", "C1")),
        (cell_id("Inputs", "B1"), cell_id("Inputs", "C1")),
    }
    names = [r for r in result["rules"] if r["kind"] == "named_range"]
    assert {tuple(r["members"]) for r in names} == {(cell_id("Inputs", "A1"),), (cell_id("Inputs", "B1"),)}


def test_nested_branch_arguments(tmp_path):
    wb = Workbook()
    wb.active["A1"] = "=IF(B1,IF(C1,D1,E1),F1)"
    path = tmp_path / "branches.xlsx"
    wb.save(path)
    rules = [r for r in analyze(path)["rules"] if r["kind"] == "conditional"]
    assert [r["arguments"] for r in rules] == [["B1", "IF(C1,D1,E1)", "F1"], ["C1", "D1", "E1"]]


def test_multiple_cycles_in_one_component(tmp_path):
    wb = Workbook()
    for key, formula in {"A1": "=B1", "B1": "=A1+C1", "C1": "=B1"}.items():
        wb.active[key] = formula
    path = tmp_path / "cycles.xlsx"
    wb.save(path)
    result = analyze(path)
    group = [cell_id("Sheet", c) for c in ("A1", "B1", "C1")]
    assert result["graph"]["cyclic_components"] == [group]
    smell = next(s for s in result["smells"] if s["kind"] == "circular_reference")
    assert smell["cells"] == group
    edges = {(e["source"], e["target"]) for e in result["graph"]["edges"]}
    assert smell["path"][0] == smell["path"][-1]
    assert all(edge in edges for edge in zip(smell["path"], smell["path"][1:]))


def test_unsupported_cell_blocks_dependents(tmp_path):
    wb = Workbook()
    wb.active["A1"] = '=INDIRECT("C1")'
    wb.active["B1"] = "=A1"
    path = tmp_path / "dynamic.xlsx"
    wb.save(path)
    result = analyze(path)
    assert {c for phase in result["plan"]["phases"] for c in phase["blocked_cells"]} == {cell_id("Sheet", "A1"), cell_id("Sheet", "B1")}


def test_unused_name_with_blank_members(tmp_path):
    wb = Workbook()
    wb.defined_names.add(DefinedName("BlankInputs", attr_text="'Sheet'!A1:A2"))
    path = tmp_path / "blank.xlsx"
    wb.save(path)
    result = analyze(path)
    assert result["graph"]["nodes"] == [cell_id("Sheet", "A1"), cell_id("Sheet", "A2")]
    assert result["rules"][0]["members"] == result["graph"]["nodes"]
    assert result["rules"][0]["id"] in result["plan"]["phases"][0]["rules"]


def test_sparse_sheet_and_dates(tmp_path):
    from datetime import datetime
    wb = Workbook()
    wb.active["A1"] = datetime(2020, 1, 1)
    wb.active["XFD1048576"] = "=A1"
    path = tmp_path / "sparse.xlsx"
    wb.save(path)
    result = analyze(path)
    assert result["summary"]["cell_count"] == 2
    assert result["inventory"]["cells"][cell_id("Sheet", "A1")]["value"] == "2020-01-01T00:00:00"


@pytest.mark.parametrize("function,formula", [
    ("HLOOKUP", "=HLOOKUP(A1,B1:C2,2,FALSE)"),
    ("XLOOKUP", "=XLOOKUP(A1,B1:B2,C1:C2)"),
])
def test_other_lookup_shapes(tmp_path, function, formula):
    wb = Workbook()
    wb.active["D1"] = formula
    path = tmp_path / "lookup.xlsx"
    wb.save(path)
    rule = next(r for r in analyze(path)["rules"] if r["kind"] == "lookup")
    assert rule["function"] == function
    assert cell_id("Sheet", "D1") in rule["cells"]
    assert cell_id("Sheet", "C2") in rule["cells"]
    assert len(rule["arguments"]) >= 3


def test_self_cycle_and_long_chain(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "=A1"
    ws["B1"] = 1
    for row in range(2, 1200):
        ws[f"B{row}"] = f"=B{row - 1}"
    wb.defined_names.add(DefinedName("OutputChain", attr_text="'Sheet'!B1199"))
    path = tmp_path / "chain.xlsx"
    wb.save(path)
    result = analyze(path)
    assert result["graph"]["cycles"] == [[cell_id("Sheet", "A1"), cell_id("Sheet", "A1")]]
    assert len(result["plan"]["phases"]) == 1199


def test_cross_process_hash_seed_determinism(models):
    code = "from sre.analysis import analyze; from sre.cli import serialize; import sys; print(serialize(analyze(sys.argv[1])), end='')"
    outputs = []
    for seed in ("1", "42"):
        outputs.append(subprocess.check_output([sys.executable, "-c", code, str(models["pricing"])], env={**os.environ, "PYTHONHASHSEED": seed}))
    assert outputs[0] == outputs[1]

import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

from sre.analysis import BANNER


def run(*args):
    return subprocess.run([sys.executable, "-m", "sre", *map(str, args)], capture_output=True, text=True, encoding="utf-8")


@pytest.mark.parametrize("command", ["parse", "graph", "rules", "smells", "plan"])
def test_cli_subcommands(models, command):
    result = run(command, models["pricing"])
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["banner"] == BANNER
    assert data["inventory" if command == "parse" else command]


@pytest.mark.parametrize("model", ["clean", "pricing"])
@pytest.mark.parametrize("format", ["md", "json"])
def test_end_to_end_cli(models, tmp_path, model, format):
    before = models[model].read_bytes()
    output = tmp_path / f"{model}.{format}"
    result = run("report", models[model], "--format", format, "--out", output)
    assert result.returncode == 0, result.stderr
    content = output.read_text(encoding="utf-8")
    assert BANNER in content
    assert models[model].read_bytes() == before
    repeat = run("report", models[model], "--format", format)
    assert repeat.stdout == content
    if format == "json":
        data = json.loads(content)
        assert data["rules"] and data["plan"]["phases"]
    else:
        assert "## Rule catalogue" in content and "## Migration plan" in content


def test_fixture_cli(tmp_path):
    result = run("fixtures", "--out", tmp_path)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["banner"] == BANNER
    assert (tmp_path / "clean.xlsx").exists() and (tmp_path / "pricing.xlsx").exists()


def test_malformed_workbook_cli(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"This is not an OOXML workbook")
    result = run("parse", bad)
    assert result.returncode == 1 and "sre:" in result.stderr
    assert "Traceback" not in result.stderr


def test_cli_protects_input_and_handles_missing_file(models, tmp_path):
    before = models["clean"].read_bytes()
    result = run("report", models["clean"], "--out", models["clean"])
    assert result.returncode != 0
    assert "must not overwrite" in result.stderr
    assert models["clean"].read_bytes() == before
    result = run("parse", tmp_path / "missing.xlsx")
    assert result.returncode == 1 and "sre:" in result.stderr


def test_cli_protects_hardlink_alias(models, tmp_path):
    alias = tmp_path / "alias.xlsx"
    os.link(models["clean"], alias)
    before = alias.read_bytes()
    result = run("report", models["clean"], "--out", alias)
    assert result.returncode != 0 and "alias" in result.stderr
    assert alias.read_bytes() == before


def test_no_em_dash_in_project():
    root = Path(__file__).resolve().parents[1]
    files = [*root.glob("sre/*.py"), *root.glob("tests/*.py"), *root.glob("*.md"), root / "pyproject.toml"]
    assert all(chr(0x2014) not in path.read_text(encoding="utf-8") for path in files)

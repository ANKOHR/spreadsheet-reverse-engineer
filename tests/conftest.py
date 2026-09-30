import pytest

from sre.fixtures import generate


@pytest.fixture
def models(tmp_path):
    generate(tmp_path)
    return {name: tmp_path / f"{name}.xlsx" for name in ("clean", "pricing")}

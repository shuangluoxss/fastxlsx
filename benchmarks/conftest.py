import pytest
from pathlib import Path


@pytest.fixture(scope="session")
def xlsx_dir():
    base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
    target_dir = base_dir / "benchmarks" / "xlsx_files"

    target_dir.mkdir(parents=True, exist_ok=True)

    return target_dir


def pytest_configure(config):
    base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()

    benchmark_defaults = {
        "benchmark_json": open(base_dir / "benchmarks" / "benchmark.json", "wb+"),
        "benchmark_columns": ["mean"],
        "benchmark_min_rounds": 10,
        "benchmark_warmup": True,
    }
    for opt, value in benchmark_defaults.items():
        setattr(config.option, opt, value)

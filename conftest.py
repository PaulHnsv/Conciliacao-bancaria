import pytest
import tempfile
import os


@pytest.fixture
def tmp_path(tmp_path_factory):
    """Override tmp_path to use /tmp on Linux to avoid permission errors on mounted Windows FS."""
    return tmp_path_factory.mktemp("test")


def pytest_configure(config):
    """Force pytest temp dir to /tmp to avoid Windows FS mount issues."""
    config.option.basetemp = "/tmp/pytest_conciliacao"

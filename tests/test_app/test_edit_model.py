import shutil

import pytest
import yaml

from edit_model import EditModel


@pytest.fixture(scope='module')
def clear_dir():
    shutil.rmtree(r'I:\rastr_add2\test_rm\РМ sect', ignore_errors=True)
    shutil.rmtree(r'I:\rastr_add2\test_rm\РМ v2', ignore_errors=True)


def test_edit_models_all(clear_dir, name_file=r'I:\rastr_add2\test_rm\test cor rm.cor'):
    with open(name_file) as f:
        print(EditModel(yaml.safe_load(f)).run())


def test_edit_model_section(clear_dir, name_file=r'I:\rastr_add2\test_rm\test section.cor'):
    with open(name_file) as f:
        print(EditModel(yaml.safe_load(f)).run())


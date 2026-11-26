from pathlib import Path
from uuid import uuid4
import pytest
from pydantic import ValidationError
from app.models import Finding
from app.workspace import Workspace


def test_discrepancy_requires_evidence():
    with pytest.raises(ValidationError):
        Finding(category='discrepancy', severity='high', description='Diferencia', evidence=[])


def test_workspace_uses_generated_ids_and_preserves_original(tmp_path: Path):
    workspace = Workspace(tmp_path)
    case_id, document_id = uuid4(), uuid4()
    path = workspace.save_original(case_id, document_id, b'original')
    assert path.read_bytes() == b'original'
    with pytest.raises(FileExistsError):
        workspace.save_original(case_id, document_id, b'replacement')
    assert path.read_bytes() == b'original'


def test_workspace_rejects_untrusted_path_identifiers(tmp_path: Path):
    with pytest.raises((ValueError, TypeError, AttributeError)):
        Workspace(tmp_path).save_original('../escape', uuid4(), b'data')

from pathlib import Path
from uuid import UUID

class Workspace:
    """Storage primitive; callers must enforce permissions, quotas and content checks."""
    def __init__(self, root: Path):
        self.root = root.resolve()

    def save_original(self, case_id: UUID, document_id: UUID, content: bytes) -> Path:
        if not isinstance(case_id, UUID) or not isinstance(document_id, UUID):
            raise TypeError("Identifiers must be UUIDs")
        directory = self.root / str(case_id) / "originals"
        directory.mkdir(parents=True, exist_ok=True)
        if not directory.resolve().is_relative_to(self.root):
            raise ValueError("Workspace outside storage root")
        path = directory / str(document_id)
        with path.open("xb") as file:
            file.write(content)
        return path

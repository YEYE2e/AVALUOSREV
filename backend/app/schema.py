import json
from .models import AuditResult

if __name__ == "__main__":
    print(json.dumps(AuditResult.model_json_schema(), ensure_ascii=False, indent=2))

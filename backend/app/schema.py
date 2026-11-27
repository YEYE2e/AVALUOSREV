import json

from .models import AgentReviewDraft, AuditResult, VerifierOutput

if __name__ == "__main__":
    print(
        json.dumps(
            {
                "AuditResult": AuditResult.model_json_schema(),
                "AgentReviewDraft": AgentReviewDraft.model_json_schema(),
                "VerifierOutput": VerifierOutput.model_json_schema(),
            },
            ensure_ascii=False,
            indent=2,
        )
    )

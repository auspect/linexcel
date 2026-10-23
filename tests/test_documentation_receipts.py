"""A reviewed card needs two distinct, complete, source-bound SDK receipts."""

import copy
import hashlib
import json

import pytest

pytest_plugins = ["test_manual_validation"]


def digest(text):
    return hashlib.sha256(text.encode()).hexdigest()


def reviewed(target, receipts):
    blob = json.dumps({"id": target, "value_source": "file_cache"})
    stages = []
    for phase in ("draft", "review"):
        system = f"Source-led {phase}"
        user = f"Source: {blob}"
        if phase == "review":
            system += " SPECIALIZED REVIEW PASS."
            user += (
                "\n\nCandidate to review (untrusted, not evidence):\n"
                + stages[0]["raw_response"]
            )
        raw = f"{target}: {phase} response grounded in supplied evidence."
        usage = {"input_tokens": 3, "output_tokens": 5, "requests": 1}
        stages.append(
            {
                "phase": phase,
                "status": "complete",
                "system_prompt": system,
                "user_prompt": user,
                "system_prompt_sha256": digest(system),
                "user_prompt_sha256": digest(user),
                "prompt_text_sha256": digest(system + "\n\n" + user),
                "raw_response": raw,
                "content_sha256": digest(raw),
                "usage": usage,
            }
        )
        receipts.append(
            {
                "response_id": f"{target}-{phase}",
                "status": "complete",
                "finish_reason": "stop",
                "prompt_text_sha256": digest(system + "\n\n" + user),
                "content_sha256": digest(raw),
                "raw_response": raw,
                "usage": copy.deepcopy(usage),
            }
        )
    return {
        "raw_response": stages[-1]["raw_response"],
        "sent_dossier": {"json_text": blob, "json_text_sha256": digest(blob)},
        "specialized_review": {
            "enabled": True,
            "status": "completed",
            "stages": stages,
        },
    }


def test_two_actual_phases_link_and_stay_distinct_from_factual_acceptance(manual):
    receipts = []
    checks = {key: reviewed(key, receipts) for key in ("workbook", "S!B2")}
    checks["request_receipts"] = receipts
    result = manual.link_documentation_receipts(checks, ["workbook", "S!B2"])
    assert result["passed"] and result["linked_stages"] == 4
    assert result["factual_certification"] is False
    assert {r["documentation_evidence"]["phase"] for r in receipts} == {
        "draft",
        "review",
    }
    assert checks["S!B2"]["raw_response"].endswith("supplied evidence.")


def test_recycled_provider_ids_preserve_distinct_local_documentation_receipts(manual):
    receipts = []
    checks = {key: reviewed(key, receipts) for key in ("workbook", "S!B2")}
    checks["request_receipts"] = receipts
    for receipt in receipts:
        receipt["response_id"] = "chatcmpl-51"
    result = manual.link_documentation_receipts(checks, ["workbook", "S!B2"])
    assert result["passed"] and result["linked_stages"] == 4
    stages = [
        stage
        for target in ("workbook", "S!B2")
        for stage in checks[target]["specialized_review"]["stages"]
    ]
    assert {stage["provider_receipt_index"] for stage in stages} == {0, 1, 2, 3}
    assert all(stage["response_id"] == "chatcmpl-51" for stage in stages)
    assert all(receipt["response_id"] == "chatcmpl-51" for receipt in receipts)


def test_one_local_receipt_cannot_cover_repeated_documentation_targets(manual):
    receipts = []
    checks = {"S!B2": reviewed("S!B2", receipts), "request_receipts": receipts}
    result = manual.link_documentation_receipts(checks, ["S!B2", "S!B2"])
    assert not result["passed"]
    assert result["linked_stages"] == 2


@pytest.mark.parametrize(
    "mutation",
    [
        "missing_review",
        "disabled",
        "incomplete",
        "wrong_final",
        "wrong_raw",
        "wrong_user",
        "wrong_system_hash",
        "length",
        "usage",
        "missing_id",
        "unassigned_receipt",
        "wrong_role",
        "missing_dossier",
        "changed_dossier",
        "wrong_combined_hash",
        "review_without_candidate",
        "review_without_role",
    ],
)
def test_phase_coverage_rejects_incomplete_or_misattributed_evidence(manual, mutation):
    receipts = []
    report = reviewed("S!B2", receipts)
    review = report["specialized_review"]
    stages = review["stages"]
    if mutation == "missing_review":
        stages.pop()
    elif mutation == "disabled":
        review["enabled"] = False
    elif mutation == "incomplete":
        review["status"] = "failed"
    elif mutation == "wrong_final":
        report["raw_response"] = stages[0]["raw_response"]
    elif mutation == "wrong_raw":
        receipts[1]["raw_response"] = "different text"
    elif mutation == "wrong_user":
        stages[1]["user_prompt"] = "No source present"
    elif mutation == "wrong_system_hash":
        stages[1]["system_prompt_sha256"] = "modified"
    elif mutation == "wrong_combined_hash":
        stages[1]["prompt_text_sha256"] = "modified"
    elif mutation in {"review_without_candidate", "review_without_role"}:
        if mutation == "review_without_candidate":
            stages[1]["user_prompt"] = "Source: " + report["sent_dossier"]["json_text"]
        else:
            stages[1]["system_prompt"] = "Source-led review"
        system = stages[1]["system_prompt"]
        user = stages[1]["user_prompt"]
        stages[1]["system_prompt_sha256"] = digest(system)
        stages[1]["user_prompt_sha256"] = digest(user)
        stages[1]["prompt_text_sha256"] = digest(system + "\n\n" + user)
        receipts[1]["prompt_text_sha256"] = stages[1]["prompt_text_sha256"]
    elif mutation == "length":
        receipts[1]["finish_reason"] = "length"
    elif mutation == "usage":
        receipts[1]["usage"]["output_tokens"] = 99
    elif mutation == "missing_id":
        receipts[1]["response_id"] = ""
    elif mutation == "unassigned_receipt":
        receipts.append(copy.deepcopy(receipts[0]))
    elif mutation == "wrong_role":
        stages[1]["phase"] = "draft"
    elif mutation == "missing_dossier":
        report.pop("sent_dossier")
    elif mutation == "changed_dossier":
        report["sent_dossier"]["json_text"] = '{"value_source":"engine"}'
    checks = {"S!B2": report, "request_receipts": receipts}
    result = manual.link_documentation_receipts(checks, ["S!B2"])
    assert not result["passed"]

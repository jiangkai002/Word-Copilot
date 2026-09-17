from app.models.edit import EditRequest, EditTargetPayload
from app.services.edit_service import _build_edit_messages


def _request(**overrides: object) -> EditRequest:
    values: dict[str, object] = {
        "conversation_id": "test-conversation",
        "instruction": "把这段话写得更清晰",
        "target": EditTargetPayload(type="paragraph", text="原文", text_hash="hash"),
        "regenerate_count": 1,
        "avoid_texts": ["上一版"],
    }
    values.update(overrides)
    return EditRequest.model_validate(values)


def test_regeneration_feedback_is_added_as_separate_requirement() -> None:
    messages = _build_edit_messages(_request(regeneration_feedback="  语气更自然  "))

    user_prompt = messages[1]["content"]
    assert "【用户修改要求】\n把这段话写得更清晰" in user_prompt
    assert "【本次重新生成的补充要求】\n语气更自然" in user_prompt
    assert "历史版本 1：上一版" in user_prompt


def test_blank_regeneration_feedback_is_omitted() -> None:
    messages = _build_edit_messages(_request(regeneration_feedback="   "))

    assert "本次重新生成的补充要求" not in messages[1]["content"]

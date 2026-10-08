"""Tests for the translation service."""

import pytest

from app.config import settings
from app.services.translation_service import TranslationError, TranslationService


class FakeResponse:
    def __init__(self, status_code, payload=None, text=""):
        self.status_code = status_code
        self._payload = payload
        self.text = text

    def json(self):
        return self._payload


def test_mock_provider_translates_both_languages():
    service = TranslationService("mock")
    assert service.translate_to_tamil(["Purchase Order"]) == {"Purchase Order": "[ta] Purchase Order"}
    assert service.translate_to_hindi(["Purchase Order"]) == {"Purchase Order": "[hi] Purchase Order"}


def test_duplicates_and_blank_texts_are_skipped():
    service = TranslationService("mock")
    result = service.translate_to_tamil(["Quantity", "Quantity", "", "   "])
    assert result == {"Quantity": "[ta] Quantity"}


def test_unknown_provider_is_rejected():
    with pytest.raises(ValueError):
        TranslationService("does_not_exist")


def test_google_cloud_requires_api_key(monkeypatch):
    monkeypatch.setattr(settings, "translation_api_key", "")
    with pytest.raises(TranslationError):
        TranslationService("google").translate_to_tamil(["Quantity"])


def test_google_cloud_request_and_response(monkeypatch):
    calls = []

    def fake_post(url, params, json, timeout):
        calls.append({"url": url, "params": params, "json": json})
        translations = [{"translatedText": f"{json['target']}:{text}"} for text in json["q"]]
        return FakeResponse(200, {"data": {"translations": translations}})

    monkeypatch.setattr(settings, "translation_api_key", "test-key")
    monkeypatch.setattr("app.services.translation_service.requests.post", fake_post)

    result = TranslationService("google").translate_to_hindi(["Quantity", "Unit Price"])

    assert result == {"Quantity": "hi:Quantity", "Unit Price": "hi:Unit Price"}
    assert calls[0]["params"] == {"key": "test-key"}
    assert calls[0]["json"]["source"] == "en"
    assert calls[0]["json"]["format"] == "text"


def test_google_cloud_error_becomes_translation_error(monkeypatch):
    monkeypatch.setattr(settings, "translation_api_key", "test-key")
    monkeypatch.setattr("app.services.translation_service.time.sleep", lambda seconds: None)
    monkeypatch.setattr(
        "app.services.translation_service.requests.post",
        lambda *args, **kwargs: FakeResponse(403, text="API key not valid"),
    )
    with pytest.raises(TranslationError):
        TranslationService("google").translate_to_tamil(["Quantity"])


def test_cloud_batches_respect_request_limits():
    service = TranslationService("mock")
    batches = service._make_batches(["text"] * 250, 100, 20000)
    assert [len(batch) for batch in batches] == [100, 100, 50]


def test_long_text_is_split_for_the_free_endpoint():
    service = TranslationService("mock")
    text = " ".join(["word"] * 3000)
    chunks = service._split_long_text(text)
    assert len(chunks) > 1
    assert all(len(chunk) <= 4500 for chunk in chunks)
    assert " ".join(chunks) == text


class FakeGoogleTranslator:
    """Stands in for deep_translator.GoogleTranslator."""

    requests = []
    drop_lines = False

    def __init__(self, source, target):
        self.target = target

    def translate(self, text):
        FakeGoogleTranslator.requests.append(text)
        lines = [f"{self.target}:{line}" for line in text.split("\n")]
        if FakeGoogleTranslator.drop_lines and len(lines) > 1:
            lines = lines[:-1]
        return "\n".join(lines)


@pytest.fixture
def fake_google(monkeypatch):
    FakeGoogleTranslator.requests = []
    FakeGoogleTranslator.drop_lines = False
    monkeypatch.setattr("deep_translator.GoogleTranslator", FakeGoogleTranslator)
    monkeypatch.setattr("app.services.translation_service.FREE_REQUEST_INTERVAL", 0)
    # The web endpoint is blocked, so the deep-translator fallback is used
    monkeypatch.setattr("app.services.translation_service.time.sleep", lambda seconds: None)
    monkeypatch.setattr(
        "app.services.translation_service.requests.post",
        lambda *args, **kwargs: FakeResponse(429),
    )
    return FakeGoogleTranslator


def test_google_free_uses_web_endpoint_first(fake_google, monkeypatch):
    calls = []

    def fake_post(url, params, data, headers, timeout):
        calls.append({"params": params, "q": data["q"]})
        return FakeResponse(200, [f"{params['tl']}:{text}" for text in data["q"]])

    monkeypatch.setattr("app.services.translation_service.requests.post", fake_post)
    texts = ["Quantity", "Line one\nLine two", " ".join(["word"] * 3000)]
    result = TranslationService("google_free").translate_to_tamil(texts)

    assert result["Quantity"] == "ta:Quantity"
    assert result["Line one\nLine two"] == "ta:Line one\nLine two"
    # The long text is sent as several chunks and joined again
    assert result[texts[2]].count("ta:") > 1
    assert calls[0]["params"] == {"client": "gtx", "sl": "en", "tl": "ta"}
    assert all(sum(len(text) for text in call["q"]) <= 4500 for call in calls)
    assert fake_google.requests == []


def test_google_free_sends_texts_in_one_batch(fake_google):
    texts = ["Quantity", "Unit Price", "Line one\nLine two"]
    result = TranslationService("google_free").translate_to_tamil(texts)

    assert result == {
        "Quantity": "ta:Quantity",
        "Unit Price": "ta:Unit Price",
        "Line one\nLine two": "ta:Line one\nta:Line two",
    }
    # One request for the two single-line texts, one for the multi-line text
    assert len(fake_google.requests) == 2


def test_google_free_falls_back_when_lines_do_not_match(fake_google):
    fake_google.drop_lines = True
    result = TranslationService("google_free").translate_to_hindi(["Quantity", "Unit Price"])

    assert result == {"Quantity": "hi:Quantity", "Unit Price": "hi:Unit Price"}
    assert len(fake_google.requests) == 3

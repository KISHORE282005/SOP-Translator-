"""
Translation Service for PO Translation System.

Translates English text segments to Tamil and Hindi through a pluggable
translation provider.
"""

import threading
import time
from typing import Callable, Dict, List

import requests

from app.config import settings
from app.utils.logger import setup_logger

logger = setup_logger(__name__)

SOURCE_LANGUAGE = "en"
LANGUAGE_CODES = {"Tamil": "ta", "Hindi": "hi"}

# Google Cloud Translation v2 request limits (kept well under the hard limits)
CLOUD_BATCH_SIZE = 100
CLOUD_BATCH_CHARS = 20000

# The public Google Translate endpoint rejects texts of 5000 characters or more
# and allows about 5 requests per second
FREE_MAX_CHARS = 4500
FREE_BATCH_SIZE = 100
FREE_REQUEST_INTERVAL = 0.4
FREE_WEB_ENDPOINT = "https://translate.googleapis.com/translate_a/t"
FREE_WEB_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)

_free_request_lock = threading.Lock()
_last_free_request = 0.0

MAX_RETRIES = 3


class TranslationError(Exception):
    """Raised when the translation provider cannot translate the content."""


class TranslationService:
    """
    Service for translating English text to Tamil and Hindi.

    The provider is selected with the TRANSLATION_PROVIDER setting:
    google_free, google or mock.
    """

    def __init__(self, provider: str = None) -> None:
        self.provider = (provider or settings.translation_provider).lower()
        providers: Dict[str, Callable[[List[str], str], List[str]]] = {
            "google_free": self._translate_google_free,
            "google": self._translate_google_cloud,
            "mock": self._translate_mock,
        }
        if self.provider not in providers:
            raise ValueError(
                f"Unknown TRANSLATION_PROVIDER '{self.provider}'. "
                f"Supported providers: {', '.join(sorted(providers))}"
            )
        self._translate_batch = providers[self.provider]

    def translate_to_tamil(self, texts: List[str]) -> Dict[str, str]:
        """Translate English texts to Tamil. Returns {english: tamil}."""
        return self.translate(texts, "Tamil")

    def translate_to_hindi(self, texts: List[str]) -> Dict[str, str]:
        """Translate English texts to Hindi. Returns {english: hindi}."""
        return self.translate(texts, "Hindi")

    def translate(self, texts: List[str], language: str) -> Dict[str, str]:
        """
        Translate a list of English texts to the given language.

        Args:
            texts: English text segments
            language: Target language name (Tamil or Hindi)

        Returns:
            Dict mapping each English text to its translation

        Raises:
            TranslationError: If the provider fails
        """
        if language not in LANGUAGE_CODES:
            raise ValueError(f"Unsupported language: {language}")

        unique_texts = list(dict.fromkeys(text for text in texts if text and text.strip()))
        if not unique_texts:
            return {}

        logger.info("Translation started", language=language,
                    provider=self.provider, text_count=len(unique_texts))
        try:
            translated = self._translate_batch(unique_texts, LANGUAGE_CODES[language])
        except TranslationError:
            raise
        except Exception as e:
            raise TranslationError(f"{language} translation failed: {e}") from e

        if len(translated) != len(unique_texts):
            raise TranslationError(
                f"{language} translation returned {len(translated)} results "
                f"for {len(unique_texts)} texts"
            )

        logger.info("Translation completed", language=language, translated_count=len(translated))
        # An empty translation would blank out the cell, so keep the English text
        return {
            source: (target if target and target.strip() else source)
            for source, target in zip(unique_texts, translated)
        }

    # Providers

    def _translate_google_cloud(self, texts: List[str], target: str) -> List[str]:
        """Translate with the Google Cloud Translation API v2 (API key required)."""
        if not settings.translation_api_key:
            raise TranslationError(
                "TRANSLATION_API_KEY is not set. Set it in .env or use "
                "TRANSLATION_PROVIDER=google_free."
            )

        results: List[str] = []
        for batch in self._make_batches(texts, CLOUD_BATCH_SIZE, CLOUD_BATCH_CHARS):
            def send(batch: List[str] = batch) -> List[str]:
                response = requests.post(
                    settings.google_translate_api_endpoint,
                    params={"key": settings.translation_api_key},
                    json={"q": batch, "source": SOURCE_LANGUAGE, "target": target, "format": "text"},
                    timeout=settings.translation_timeout,
                )
                if response.status_code != 200:
                    raise TranslationError(
                        f"Google Translation API returned HTTP {response.status_code}: "
                        f"{response.text[:200]}"
                    )
                translations = response.json()["data"]["translations"]
                return [item["translatedText"] for item in translations]

            results.extend(self._with_retries(send))
        return results

    def _translate_google_free(self, texts: List[str], target: str) -> List[str]:
        """
        Translate with the public Google Translate endpoints (no API key).

        Google rate limits each public endpoint separately, so when the first
        one is blocked the second one usually still works.
        """
        try:
            return self._translate_google_web(texts, target)
        except TranslationError as e:
            logger.warning("Google web endpoint failed, falling back to deep-translator",
                           error=str(e))
            return self._translate_deep_translator(texts, target)

    def _translate_google_web(self, texts: List[str], target: str) -> List[str]:
        """Translate with the translate.googleapis.com endpoint, many texts per request."""
        # Long texts are sent as several chunks and joined again afterwards
        chunks: List[str] = []
        chunk_counts: List[int] = []
        for text in texts:
            parts = self._split_long_text(text)
            chunks.extend(parts)
            chunk_counts.append(len(parts))

        translated: List[str] = []
        for batch in self._make_batches(chunks, FREE_BATCH_SIZE, FREE_MAX_CHARS):
            def send(batch: List[str] = batch) -> List[str]:
                self._wait_for_free_slot()
                response = requests.post(
                    FREE_WEB_ENDPOINT,
                    params={"client": "gtx", "sl": SOURCE_LANGUAGE, "tl": target},
                    data={"q": batch},
                    headers={"User-Agent": FREE_WEB_USER_AGENT},
                    timeout=settings.translation_timeout,
                )
                if response.status_code != 200:
                    raise TranslationError(
                        f"Google Translate web endpoint returned HTTP {response.status_code}"
                    )
                # Each item is the translation, or [translation, detected language]
                items = [item[0] if isinstance(item, list) else item for item in response.json()]
                if len(items) != len(batch):
                    raise TranslationError(
                        f"Google Translate web endpoint returned {len(items)} results "
                        f"for {len(batch)} texts"
                    )
                return items

            translated.extend(self._with_retries(send))

        results: List[str] = []
        position = 0
        for count in chunk_counts:
            results.append(" ".join(translated[position:position + count]))
            position += count
        return results

    def _translate_deep_translator(self, texts: List[str], target: str) -> List[str]:
        """
        Translate with the translate.google.com mobile page through deep-translator.

        The endpoint allows only a few requests per second, so many texts are
        sent together in one request, one text per line.
        """
        try:
            from deep_translator import GoogleTranslator
        except ImportError as e:
            raise TranslationError(
                "The google_free provider needs the 'deep-translator' package. "
                "Run: pip install deep-translator"
            ) from e

        def send(text: str) -> str:
            def request() -> str:
                self._wait_for_free_slot()
                return GoogleTranslator(source=SOURCE_LANGUAGE, target=target).translate(text)

            return self._with_retries(request) or ""

        def translate_one(text: str) -> str:
            return " ".join(send(chunk) or chunk for chunk in self._split_long_text(text))

        results: Dict[str, str] = {}
        single_line = [text for text in texts if "\n" not in text and len(text) <= FREE_MAX_CHARS]
        for batch in self._make_batches(single_line, FREE_BATCH_SIZE, FREE_MAX_CHARS, separator_chars=1):
            lines = send("\n".join(batch)).split("\n")
            if len(lines) == len(batch):
                results.update(zip(batch, (line.strip() for line in lines)))
            else:
                # The line structure was lost, so the texts cannot be matched safely
                logger.warning("Batch translation returned a different line count, "
                               "translating texts one by one", batch_size=len(batch))

        return [results[text] if text in results else translate_one(text) for text in texts]

    def _translate_mock(self, texts: List[str], target: str) -> List[str]:
        """Fake offline translation, used for tests and demos."""
        return [f"[{target}] {text}" for text in texts]

    # Helpers

    def _with_retries(self, func: Callable[[], object]):
        """Call func, retrying with a short backoff on failure."""
        last_error: Exception = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return func()
            except Exception as e:
                last_error = e
                logger.warning("Translation request failed", attempt=attempt, error=str(e))
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 2)
        raise TranslationError(f"Translation provider '{self.provider}' failed: {last_error}")

    def _wait_for_free_slot(self) -> None:
        """Keep requests to the free endpoint below its rate limit, across all threads."""
        global _last_free_request
        with _free_request_lock:
            wait = _last_free_request + FREE_REQUEST_INTERVAL - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            _last_free_request = time.monotonic()

    def _make_batches(self, texts: List[str], max_items: int, max_chars: int,
                      separator_chars: int = 0) -> List[List[str]]:
        """Group texts into batches that respect the provider's request limits."""
        batches: List[List[str]] = []
        current: List[str] = []
        current_chars = 0
        for text in texts:
            size = len(text) + separator_chars
            if current and (len(current) >= max_items or current_chars + size > max_chars):
                batches.append(current)
                current, current_chars = [], 0
            current.append(text)
            current_chars += size
        if current:
            batches.append(current)
        return batches

    def _split_long_text(self, text: str) -> List[str]:
        """Split a text into chunks small enough for the free endpoint."""
        if len(text) <= FREE_MAX_CHARS:
            return [text]

        chunks: List[str] = []
        current = ""
        for word in text.split(" "):
            # A single word longer than the limit has to be cut hard
            while len(word) > FREE_MAX_CHARS:
                if current:
                    chunks.append(current)
                    current = ""
                chunks.append(word[:FREE_MAX_CHARS])
                word = word[FREE_MAX_CHARS:]
            candidate = f"{current} {word}" if current else word
            if len(candidate) > FREE_MAX_CHARS:
                chunks.append(current)
                current = word
            else:
                current = candidate
        if current:
            chunks.append(current)
        return chunks

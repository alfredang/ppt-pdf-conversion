"""Chinese-to-English translation module using deep-translator."""

import time
from deep_translator import GoogleTranslator


def has_cjk(text):
    """Check if text contains CJK (Chinese/Japanese/Korean) characters."""
    return any(
        '\u4e00' <= c <= '\u9fff' or  # CJK Unified Ideographs
        '\u3400' <= c <= '\u4dbf' or  # CJK Extension A
        '\u3000' <= c <= '\u303f' or  # CJK Symbols and Punctuation
        '\uff00' <= c <= '\uffef'     # Fullwidth Forms
        for c in text
    )


class Translator:
    def __init__(self):
        self.translator = GoogleTranslator(source='zh-TW', target='en')
        self.cache = {}
        self.last_request_time = 0
        self.min_delay = 0.3  # seconds between requests

    def _rate_limit(self):
        """Enforce minimum delay between translation requests."""
        elapsed = time.time() - self.last_request_time
        if elapsed < self.min_delay:
            time.sleep(self.min_delay - elapsed)
        self.last_request_time = time.time()

    def translate_text(self, text):
        """Translate a single text string from Chinese to English."""
        if not text or not text.strip():
            return text

        if not has_cjk(text):
            return text

        # Check cache
        if text in self.cache:
            return self.cache[text]

        try:
            self._rate_limit()
            result = self.translator.translate(text)
            if result:
                self.cache[text] = result
                return result
        except Exception as e:
            print(f"Translation error for '{text[:30]}...': {e}")

        return text  # Return original on failure

    def translate_page(self, page_data):
        """Translate all text groups in a page."""
        for group in page_data.get('text_groups', []):
            original_text = group['text']
            translated = self.translate_text(original_text)
            group['translated_text'] = translated

            # Also translate individual spans for reference
            for span in group.get('spans', []):
                span['translated_text'] = self.translate_text(span['text'])

        return page_data

    def translate_pages(self, pages_data, progress_callback=None):
        """Translate all pages."""
        for i, page_data in enumerate(pages_data):
            if progress_callback:
                progress_callback(f"Translating page {i + 1}/{len(pages_data)}")
            self.translate_page(page_data)
        return pages_data

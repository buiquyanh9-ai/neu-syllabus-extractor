"""
translator.py — Dịch course_description_vi → course_description_en
bằng OpenAI API (gpt-4o-mini). Chỉ gọi khi --translate được bật.

Chi phí ước tính: ~$0.0005 / mô tả (gpt-4o-mini, ~400 tokens input)
→ 4909 file ≈ $2.5 tổng cộng

Độ chính xác: tốt cho văn bản học thuật tiếng Việt, nhưng thuật ngữ
chuyên ngành đôi khi cần review thủ công.
"""
import os
import logging
from typing import Optional

log = logging.getLogger(__name__)


def translate_vi_to_en(text: str, course_name: str = "") -> Optional[str]:
    """
    Dịch đoạn văn tiếng Việt sang tiếng Anh bằng gpt-4o-mini.
    Trả về None nếu không có key hoặc gặp lỗi.
    """
    if not text or not text.strip():
        return None

    api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        log.warning("OPENAI_API_KEY không có — bỏ qua dịch")
        return None

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        context = f' (course: "{course_name}")' if course_name else ""
        prompt  = (
            f"Translate this Vietnamese university course description to English{context}. "
            "Keep academic terminology precise. Return only the translated text, no explanation.\n\n"
            f"{text}"
        )

        response = client.chat.completions.create(
            model       = os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages    = [{"role": "user", "content": prompt}],
            max_tokens  = 1000,
            temperature = 0.2,
        )
        return response.choices[0].message.content.strip()

    except Exception as e:
        log.error(f"Translation error: {e}")
        return None

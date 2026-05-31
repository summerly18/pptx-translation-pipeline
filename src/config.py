"""Shared configuration for the PPTX translation pipeline."""

import os

__version__ = "1.0.0"

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

INPUT_DIR = os.path.join(BASE_DIR, "input")
EXTRACTED_DIR = os.path.join(BASE_DIR, "1_extracted")
TRANSLATED_DIR = os.path.join(BASE_DIR, "2_translated")
REVIEWED_DIR = os.path.join(BASE_DIR, "3_reviewed")
FINAL_DIR = os.path.join(BASE_DIR, "4_final")
GLOSSARY_DIR = os.path.join(BASE_DIR, "glossary")
GLOSSARY_FILE = os.path.join(GLOSSARY_DIR, "glossary_FR_KO.xlsx")

for directory in (
    INPUT_DIR,
    EXTRACTED_DIR,
    TRANSLATED_DIR,
    REVIEWED_DIR,
    FINAL_DIR,
    GLOSSARY_DIR,
):
    os.makedirs(directory, exist_ok=True)

LLAMA_SERVER_URL = "http://localhost:8080/completion"
LLAMA_MODEL = "qwen"
MAX_TOKENS = 1024
TEMPERATURE = 0.1
REQUEST_TIMEOUT = 120

EXCEL_HEADERS = [
    "슬라이드",
    "Shape ID",
    "Shape 이름",
    "유형",
    "원문 (FR)",
    "1차 번역 (KO)",
    "Claude 감수",
    "최종 번역",
    "매칭 용어",
    "오류 플래그",
]

FLAG_SHORT_TEXT = 3
FLAG_PATTERNS = [
    r"\d+[\.,]\d+",
]

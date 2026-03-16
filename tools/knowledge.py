"""知识库管理模块"""
import json
import os

KNOWLEDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "knowledge")

KB_FILE = {
    "field":     "fields.json",
    "pattern":   "patterns.json",
    "diagnosis": "diagnoses.json",
}

FIELD_DB = {}
PATTERNS = {}
DIAGNOSES = {}

def _load(filename: str) -> dict:
    path = os.path.join(KNOWLEDGE_DIR, filename)
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def _save(filename: str, data: dict) -> None:
    path = os.path.join(KNOWLEDGE_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def reload_knowledge():
    global FIELD_DB, PATTERNS, DIAGNOSES
    FIELD_DB = _load("fields.json")
    PATTERNS = _load("patterns.json")
    DIAGNOSES = _load("diagnoses.json")
    return FIELD_DB, PATTERNS, DIAGNOSES

def get_pattern_keys():
    return sorted(PATTERNS.keys())

reload_knowledge()

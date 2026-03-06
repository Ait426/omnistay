"""
config_loader.py — /config/config.json 로더

모든 설정값은 이 모듈을 통해서만 접근한다.
하드코딩 금지 (CLAUDE.md Rule 1).
"""

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "config.json"
_config_cache: dict | None = None


def load_config() -> dict:
    """config.json을 읽어 딕셔너리로 반환한다. 실패 시 즉시 중단(Fail-Fast)."""
    global _config_cache
    if _config_cache is not None:
        return _config_cache

    if not _CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"설정 파일을 찾을 수 없습니다: {_CONFIG_PATH}"
        )

    try:
        text = _CONFIG_PATH.read_text(encoding="utf-8")
        _config_cache = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"config.json 파싱 실패: {exc}"
        ) from exc

    logger.info("config.json 로드 완료: %s", _CONFIG_PATH)
    return _config_cache


def get(key: str, *, default=None):
    """점(.) 구분 키로 중첩 값을 조회한다. 예: get('database.url')"""
    cfg = load_config()
    parts = key.split(".")
    node = cfg
    for part in parts:
        if not isinstance(node, dict) or part not in node:
            if default is not None:
                return default
            raise KeyError(f"config.json에 '{key}' 키가 없습니다.")
        node = node[part]
    return node

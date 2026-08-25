# -*- coding: utf-8 -*-
"""
待编码.xlsx：一级开放编码专用脚本

默认输入：/Users/wueve/Desktop/待编码.xlsx
默认手册：/Users/wueve/Desktop/Essay/XL/251130计算扎根理论/step1.人工开放编码/前置结果/人工开放编码-参考.xlsx
默认输出：/Users/wueve/Desktop/一级编码输出/待编码_一级编码结果.xlsx

本脚本执行完整流程：Stage 1 一级开放编码与待搜索对象识别 →
Stage 2 豆包联网搜索 → Stage 3 基于检索证据复核并回填最终一级编码。
采用批次流水线：每完成一批待搜索识别，立即搜索该批，保存断点后再处理下一批。
它会保留输入表的全部行，并使用“原Excel行号”作为稳定记录ID；三个阶段均支持断点续跑。

安装依赖：
    pip install pandas openpyxl certifi

.env 必须配置（密钥仅保存在你本地，不写入脚本）：
    GEMINI_API_KEY=你的新API_KEY

Stage 2 豆包联网搜索必须配置：
    ARK_API_KEY=你的火山方舟API_KEY
    ARK_MODEL=支持Responses API和web_search的模型或推理接入点

可选配置：
    GEMINI_API_URL=https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent
    GEMINI_MODEL=gemini-3.1-pro-preview

建议先测试20条：
    python 一级开放编码_待编码.py --limit 20

确认输出正常后继续全量运行：
    python 一级开放编码_待编码.py

脚本支持断点续跑。再次运行时会读取 checkpoint，不会重复请求已经成功的记录。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import re
import ssl
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen

import pandas as pd


def load_simple_dotenv() -> None:
    """轻量读取.env，避免额外依赖python-dotenv。"""
    candidates = [Path.cwd() / ".env", Path(__file__).resolve().parent / ".env"]
    for path in candidates:
        if not path.exists():
            continue
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, value)


load_simple_dotenv()


# ============================================================
# 默认配置：不传命令行参数也可以直接运行
# ============================================================

DEFAULT_INPUT_XLSX = "/Users/wueve/Desktop/编码.xlsx"
DEFAULT_CODING_BOOK_XLSX = "/Users/wueve/Desktop/Essay/XL/251130计算扎根理论/step1.人工开放编码/前置结果/人工开放编码-参考.xlsx"
DEFAULT_OUTPUT_DIR = "/Users/wueve/Desktop/一级编码输出2"

DEFAULT_API_URL = os.getenv(
    "GEMINI_API_URL",
    "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
)
DEFAULT_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.1-pro-preview",
)

TEXT_COLUMN = "原文(OT)"
ID_COLUMN = "原Excel行号"
SEQ_COLUMN = "序号"

BOOK_SENTENCE_COLUMN = "句子"
BOOK_CODE_COLUMN = "人工整合开放编码"
BOOK_NOTE_COLUMN = "备注"

# 从“共享字符串未恢复/空白”中成功找回真实原文的53条记录。
# 仅在命令行传入 --rerun-restored 时强制重跑；完成一次后应移除该参数。
RESTORED_RECORD_IDS = frozenset({
    "154472", "334360", "334729", "336345", "336918", "337835", "338036", "339275",
    "344916", "345249", "347270", "347413", "347442", "347514", "347857", "348188",
    "349610", "350245", "351202", "351876", "352026", "352064", "352137", "357677",
    "358236", "358392", "359068", "359402", "360299", "363981", "365281", "367006",
    "370169", "370259", "371522", "372195", "372384", "372877", "373305", "375004",
    "375170", "375513", "375682", "376185", "376434", "376473", "376482", "376536",
    "376996", "377326", "377887", "377992", "378170",
})


# ============================================================
# 通用工具
# ============================================================

def normalize_text(value: Any) -> str:
    """只做空格标准化，不删除网络语言、表情或标点。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\u3000", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_record_id(value: Any) -> str:
    """把Excel行号转换为稳定字符串，避免 3288.0 之类的变化。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def clamp_confidence(value: Any, default: float = 0.70) -> float:
    try:
        score = float(value)
        if score > 1:
            score /= 100.0
        return round(max(0.0, min(1.0, score)), 4)
    except Exception:
        return default


def parse_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "y", "是"}:
        return True
    if text in {"false", "0", "no", "n", "否", ""}:
        return False
    return default


def retry_sleep(attempt: int) -> None:
    delay = 1.5 * (2 ** attempt) * random.uniform(0.85, 1.15)
    time.sleep(min(delay, 30.0))


_GEMINI_SSL_CONTEXT: ssl.SSLContext | None = None


def get_gemini_ssl_context() -> ssl.SSLContext:
    """使用certifi的Mozilla CA证书链，兼容macOS独立安装版Python。"""
    global _GEMINI_SSL_CONTEXT
    if _GEMINI_SSL_CONTEXT is not None:
        return _GEMINI_SSL_CONTEXT
    try:
        import certifi
    except ImportError as exc:
        raise RuntimeError(
            "当前Python环境缺少certifi。请在终端运行：\n"
            "python -m pip install -U certifi"
        ) from exc
    _GEMINI_SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
    return _GEMINI_SSL_CONTEXT


def atomic_write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temp_path, path)


def load_checkpoint(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        results = data.get("results", {}) if isinstance(data, dict) else {}
        return results if isinstance(results, dict) else {}
    except Exception as exc:
        backup = path.with_suffix(path.suffix + f".broken_{int(time.time())}")
        path.rename(backup)
        print(f"⚠️ checkpoint 无法读取，已改名备份：{backup}\n错误：{exc}")
        return {}


def invalidate_changed_original_text(
    result_map: dict[str, dict],
    records: list[dict],
    label: str,
) -> list[str]:
    """输入原文被修复后，自动丢弃该记录的旧结果，避免恢复错误checkpoint。"""
    invalidated: list[str] = []
    for record in records:
        record_id = record["record_id"]
        previous = result_map.get(record_id)
        if not isinstance(previous, dict) or "original_text" not in previous:
            continue
        if normalize_text(previous.get("original_text")) == record["original_text"]:
            continue
        result_map.pop(record_id, None)
        invalidated.append(record_id)
    if invalidated:
        print(
            f"[{label}] 检测到{len(invalidated)}条原文已变化，"
            "已自动作废对应旧checkpoint并重新处理。",
            flush=True,
        )
        preview = "、".join(invalidated[:20])
        suffix = "……" if len(invalidated) > 20 else ""
        print(f"[{label}] 重新处理ID：{preview}{suffix}", flush=True)
    return invalidated


def is_unrecovered_placeholder(text: Any) -> bool:
    """识别原文恢复失败占位符，防止把系统提示当研究语料编码。"""
    value = normalize_text(text)
    return (
        "共享字符串索引" in value
        and ("未恢复" in value or "从原文件回查" in value)
    )


def reject_unrecovered_placeholders(records: list[dict]) -> None:
    bad_ids = [
        record["record_id"]
        for record in records
        if is_unrecovered_placeholder(record["original_text"])
    ]
    if not bad_ids:
        return
    preview = "、".join(bad_ids[:20])
    suffix = "……" if len(bad_ids) > 20 else ""
    raise RuntimeError(
        f"输入表仍有{len(bad_ids)}条‘共享字符串未恢复’占位记录："
        f"{preview}{suffix}\n"
        "请使用《待编码_最终核对版.xlsx》，不要继续使用旧的待编码表。"
    )


def force_remove_record_ids(
    result_map: dict[str, dict],
    record_ids: set[str] | frozenset[str],
    label: str,
) -> int:
    """一次性强制清除指定ID的阶段结果，用于完整重跑恢复记录。"""
    removed = 0
    for record_id in record_ids:
        if record_id in result_map:
            result_map.pop(record_id, None)
            removed += 1
    print(
        f"[{label}] --rerun-restored已启用："
        f"目标={len(record_ids)}条，清除旧结果={removed}条。",
        flush=True,
    )
    return removed


def force_remove_search_cache_ids(
    search_cache: dict[str, dict],
    record_ids: set[str] | frozenset[str],
) -> int:
    """删除恢复记录对应的豆包缓存，确保需要搜索的记录重新联网。"""
    remove_keys = []
    for key, item in search_cache.items():
        if not isinstance(item, dict):
            continue
        record_id = normalize_record_id(item.get("record_id"))
        if record_id in record_ids:
            remove_keys.append(key)
    for key in remove_keys:
        search_cache.pop(key, None)
    print(
        f"[豆包缓存] --rerun-restored已启用：清除旧检索缓存={len(remove_keys)}项。",
        flush=True,
    )
    return len(remove_keys)


def strip_code_fences(text: str) -> str:
    value = (text or "").strip()
    value = re.sub(r"<think>.*?</think>", "", value, flags=re.I | re.S).strip()
    value = re.sub(r"^```(?:json)?\s*", "", value, flags=re.I)
    value = re.sub(r"\s*```$", "", value)
    return value.strip()


def extract_json(text: str) -> dict:
    """从可能包含额外文字的响应中提取首个JSON对象。"""
    cleaned = strip_code_fences(text)
    decoder = json.JSONDecoder()
    for index, char in enumerate(cleaned):
        if char != "{":
            continue
        try:
            obj, _ = decoder.raw_decode(cleaned[index:])
            if isinstance(obj, dict):
                return obj
        except Exception:
            continue
    raise ValueError("模型响应中没有可解析的JSON对象")


def safe_filename(value: str) -> str:
    value = re.sub(r"[^0-9A-Za-z._\-\u4e00-\u9fff]+", "_", value)
    return value.strip("_") or "model"


# ============================================================
# 输入与输出
# ============================================================

def load_pending_records(input_path: Path, sheet: str | int = 0) -> tuple[pd.DataFrame, list[dict]]:
    df = pd.read_excel(input_path, sheet_name=sheet, dtype=object)
    required = [SEQ_COLUMN, ID_COLUMN, TEXT_COLUMN]
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(
            f"待编码文件缺少列：{missing}\n当前列名：{list(df.columns)}"
        )

    df = df.copy()
    df["_record_id"] = df[ID_COLUMN].map(normalize_record_id)
    df["_original_text"] = df[TEXT_COLUMN].map(normalize_text)

    if (df["_record_id"] == "").any():
        bad_rows = (df.index[df["_record_id"] == ""] + 2).tolist()
        raise ValueError(f"以下Excel行的“{ID_COLUMN}”为空：{bad_rows[:20]}")

    duplicated = df[df["_record_id"].duplicated(keep=False)]["_record_id"].tolist()
    if duplicated:
        raise ValueError(f"“{ID_COLUMN}”存在重复值：{duplicated[:20]}")

    records = []
    for index, row in df.iterrows():
        records.append(
            {
                "record_id": row["_record_id"],
                "sequence": row[SEQ_COLUMN],
                "source_excel_row": row[ID_COLUMN],
                "original_text": row["_original_text"],
                "input_order": int(index),
            }
        )
    return df, records


def char_ngrams(text: str, n: int = 2) -> set[str]:
    """为中文短文本生成字符n-gram，用于从编码手册检索相似案例。"""
    compact = re.sub(r"\s+", "", normalize_text(text))
    # 年份和孤立数字很容易造成伪相似，例如“2024新春档”错误匹配“总结2024”。
    compact = re.sub(r"\d+", "", compact)
    compact = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", compact)
    if not compact:
        return set()
    if len(compact) < n:
        return {compact}
    grams = {compact[i:i + n] for i in range(len(compact) - n + 1)}
    generic_grams = {
        "怎么", "么样", "什么", "哪个", "时候", "可以", "能不", "不能",
        "觉得", "你觉", "请问", "知道", "告诉", "一下", "下我", "一个", "这个",
        "是不", "不是", "是否", "我想", "我要", "问你", "有没", "没有",
    }
    return grams - generic_grams


def load_coding_book(coding_book_path: Path, sheet: str | int = 0) -> tuple[list[dict], set[str]]:
    book = pd.read_excel(coding_book_path, sheet_name=sheet, dtype=object)
    required = [BOOK_SENTENCE_COLUMN, BOOK_CODE_COLUMN]
    missing = [column for column in required if column not in book.columns]
    if missing:
        raise ValueError(
            f"编码手册缺少列：{missing}\n当前列名：{list(book.columns)}"
        )

    if BOOK_NOTE_COLUMN not in book.columns:
        book[BOOK_NOTE_COLUMN] = ""

    examples = []
    code_set: set[str] = set()
    invalid_values = {"", "/", "无", "none", "nan"}
    for _, row in book.iterrows():
        sentence = normalize_text(row.get(BOOK_SENTENCE_COLUMN))
        code = normalize_text(row.get(BOOK_CODE_COLUMN))
        note = normalize_text(row.get(BOOK_NOTE_COLUMN))
        if not sentence or code.lower() in invalid_values:
            continue
        examples.append(
            {
                "sentence": sentence,
                "code": code,
                "note": note,
                "ngrams": char_ngrams(sentence),
            }
        )
        code_set.add(code)

    if not examples:
        raise ValueError("编码手册中没有可用的人工整合开放编码示例")
    return examples, code_set


def retrieve_reference_examples(
    text: str,
    coding_book: list[dict],
    top_k: int = 3,
    min_score: float = 0.16,
) -> list[dict]:
    """只把最相近的少量人工案例放入提示词，避免750个编码全部塞入上下文。"""
    target = char_ngrams(text)
    if not target:
        return []

    scored = []
    for item in coding_book:
        candidate = item["ngrams"]
        if not candidate:
            continue
        intersection = len(target & candidate)
        if intersection == 0:
            continue
        union = len(target | candidate)
        jaccard = intersection / union if union else 0.0
        containment = intersection / min(len(target), len(candidate))
        score = 0.65 * jaccard + 0.35 * containment
        scored.append((score, item))

    scored.sort(key=lambda pair: pair[0], reverse=True)
    references = []
    used_codes = set()
    for score, item in scored:
        if score < min_score:
            break
        # 同一编码最多展示一个案例，增加参考的概念多样性。
        if item["code"] in used_codes:
            continue
        references.append(
            {
                "参考原文": item["sentence"],
                "人工一级编码": item["code"],
                "备注": item["note"],
                "相似度": round(score, 4),
            }
        )
        used_codes.add(item["code"])
        if len(references) >= top_k:
            break
    return references


def make_blank_result(record: dict) -> dict:
    return {
        "record_id": record["record_id"],
        "first_level_code": "",
        "is_new_code": "",
        "evidence": "",
        "memo": "",
        "confidence": "",
        "needs_review": True,
        "review_reason": "原文为空，未调用API；请回查原始数据。",
        "status": "blank_text",
        "error": "",
    }


def make_failed_result(record: dict, error: str) -> dict:
    """API失败时明确留空，禁止伪造“文本主旨待细化”。"""
    return {
        "record_id": record["record_id"],
        "first_level_code": "",
        "is_new_code": "",
        "evidence": "",
        "memo": "",
        "confidence": "",
        "needs_review": True,
        "review_reason": "API调用或结构解析失败，需要重新运行。",
        "status": "api_failed",
        "error": str(error)[:1000],
    }


def export_excel(
    source_df: pd.DataFrame,
    result_map: dict[str, dict],
    output_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result_rows = []
    for _, row in source_df.iterrows():
        record_id = row["_record_id"]
        result = result_map.get(record_id, {})
        result_rows.append(
            {
                SEQ_COLUMN: row[SEQ_COLUMN],
                ID_COLUMN: row[ID_COLUMN],
                TEXT_COLUMN: row[TEXT_COLUMN] if not pd.isna(row[TEXT_COLUMN]) else "",
                "一级编码": result.get("first_level_code", ""),
                "是否新增一级编码": result.get("is_new_code", ""),
                "编码证据": result.get("evidence", ""),
                "编码备忘录": result.get("memo", ""),
                "置信度": result.get("confidence", ""),
                "是否需人工复核": result.get("needs_review", ""),
                "复核原因": result.get("review_reason", ""),
                "处理状态": result.get("status", "pending"),
                "错误信息": result.get("error", ""),
            }
        )

    output_df = pd.DataFrame(result_rows)
    temp_path = output_path.with_name(output_path.stem + "_tmp.xlsx")
    output_df.to_excel(temp_path, index=False)
    os.replace(temp_path, output_path)


# ============================================================
# 一级开放编码提示词与API调用
# ============================================================

SYSTEM_PROMPT = """
你是一名严格的计算扎根理论研究助理。研究问题是：用户为什么会主动在互联网上寻找AI Agent并向其发起互动？

你的任务仅是进行“一级开放编码”，不是二级聚类，也不是三级心理需求归纳。

一级编码要求：
1. 贴近原文，概括用户在这条文本中正在做什么、向AI表达什么或期待AI做什么。
2. 每条文本只给一个主导一级编码；如同时有多个行为，选择与主动互动目的最相关的行为。
3. 编码应是完整、清楚的一句话，一般15—45个汉字。
4. 使用“向AI……”“请求AI……”“与AI分享……”等行为性表达。
5. 不要直接使用“关系回应需求、陪伴共在需求、情绪支持需求”等三级心理需求名称。
6. 输入中会提供从人工编码手册检索出的相似案例。这些案例只用于保持命名方式和编码粒度一致，不是强制分类答案。
7. 如果文本含义与某个人工案例实质相同，可以复用其“人工一级编码”，并将 is_new_code 设为 false。
8. 如果现有案例不能准确表达该文本，可以形成新的一级编码，并将 is_new_code 设为 true。不得为了复用手册而牺牲原文含义。
9. 不要把不同主题机械归为饮食建议；必须依据实际原文判断对象和目的。
10. 不确定时仍按字面谨慎概括，并将 needs_review 设为 true，说明原因。
11. evidence 必须逐字来自原文，不得改写或编造。
12. 不得输出“文本主旨待细化”作为保底编码。
13. 只输出严格合法JSON，不输出JSON之外的文字。

输出格式：
{
  "results": [
    {
      "record_id": "原样返回输入ID",
      "first_level_code": "一个一级开放编码",
      "is_new_code": false,
      "evidence": "原文中的关键短语",
      "memo": "一句话说明为什么这样编码",
      "confidence": 0.0,
      "needs_review": false,
      "review_reason": ""
    }
  ]
}
""".strip()


def build_user_prompt(records: list[dict]) -> str:
    payload = [
        {
            "record_id": record["record_id"],
            "original_text": record["original_text"],
            "人工编码手册相似案例": record.get("reference_examples", []),
        }
        for record in records
    ]
    return (
        "请对下面每条微博文本进行一级开放编码。每个record_id必须且只能返回一次。\n\n"
        + json.dumps(payload, ensure_ascii=False)
    )


def resolve_gemini_generate_url(api_url: str, model: str) -> str:
    """生成Google Gemini原生generateContent地址。"""
    if "{model}" in api_url:
        return api_url.format(model=quote(model, safe=""))
    return api_url


def test_gemini_connection(api_key: str, model: str, timeout: int = 20) -> dict:
    """启动前验证Google密钥、模型名称和generateContent权限。"""
    model_url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + quote(model, safe="")
    )
    request = Request(
        model_url,
        headers={
            "x-goog-api-key": api_key,
            "Accept": "application/json",
        },
        method="GET",
    )
    try:
        with urlopen(
            request,
            timeout=timeout,
            context=get_gemini_ssl_context(),
        ) as response:
            response_text = response.read().decode("utf-8", "replace")
    except HTTPError as exc:
        error_text = exc.read().decode("utf-8", "replace")
        raise RuntimeError(
            f"Gemini连接测试失败（HTTP {exc.code}）：{error_text[:1200]}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(f"无法连接Google Gemini API：{exc}") from exc

    data = json.loads(response_text)
    methods = data.get("supportedGenerationMethods") or []
    if "generateContent" not in methods:
        raise RuntimeError(
            f"模型{model}不支持generateContent；Google返回的方法：{methods}"
        )
    return data


def call_gemini_generate_content(
    api_url: str,
    api_key: str,
    model: str,
    records: list[dict],
    temperature: float,
    max_tokens: int,
    timeout: int,
) -> str:
    """调用Google Gemini原生generateContent REST API。"""
    headers = {
        "x-goog-api-key": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "system_instruction": {
            "parts": [{"text": SYSTEM_PROMPT}],
        },
        "contents": [
            {
                "role": "user",
                "parts": [{"text": build_user_prompt(records)}],
            }
        ],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_tokens,
            "responseMimeType": "application/json",
        },
    }

    request_body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request_url = resolve_gemini_generate_url(api_url, model)
    http_request = Request(request_url, data=request_body, headers=headers, method="POST")
    try:
        with urlopen(
            http_request,
            timeout=timeout,
            context=get_gemini_ssl_context(),
        ) as response:
            status = response.getcode()
            response_text = response.read().decode("utf-8", "replace")
    except HTTPError as exc:
        error_text = exc.read().decode("utf-8", "replace")
        raise RuntimeError(f"HTTP {exc.code}: {error_text[:1500]}") from exc
    except URLError as exc:
        raise RuntimeError(f"网络请求失败：{exc}") from exc

    if status != 200:
        raise RuntimeError(f"HTTP {status}: {response_text[:1500]}")

    data = json.loads(response_text)
    candidates = data.get("candidates") or []
    if not candidates:
        prompt_feedback = data.get("promptFeedback") or {}
        raise RuntimeError(
            "Gemini响应中没有candidates："
            + json.dumps(prompt_feedback or data, ensure_ascii=False)[:1500]
        )

    candidate = candidates[0] if isinstance(candidates[0], dict) else {}
    candidate_content = candidate.get("content") or {}
    parts = candidate_content.get("parts") or []
    text_parts = [
        part.get("text", "")
        for part in parts
        if isinstance(part, dict) and isinstance(part.get("text"), str)
    ]
    content = "".join(text_parts).strip()
    if not content:
        finish_reason = candidate.get("finishReason", "")
        raise RuntimeError(
            f"Gemini响应没有文本，finishReason={finish_reason}："
            + json.dumps(data, ensure_ascii=False)[:1500]
        )
    return content


def call_gemini_json_prompt(
    api_url: str,
    api_key: str,
    model: str,
    system_prompt: str,
    user_prompt: str,
    temperature: float,
    max_tokens: int,
    timeout: int,
) -> dict:
    """通用Gemini JSON调用，供Stage 2术语识别和Stage 3复核使用。"""
    headers = {
        "x-goog-api-key": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "system_instruction": {"parts": [{"text": system_prompt}]},
        "contents": [
            {"role": "user", "parts": [{"text": user_prompt}]},
        ],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_tokens,
            "responseMimeType": "application/json",
        },
    }
    request = Request(
        resolve_gemini_generate_url(api_url, model),
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urlopen(
            request,
            timeout=timeout,
            context=get_gemini_ssl_context(),
        ) as response:
            response_text = response.read().decode("utf-8", "replace")
    except HTTPError as exc:
        error_text = exc.read().decode("utf-8", "replace")
        raise RuntimeError(f"Gemini HTTP {exc.code}: {error_text[:1500]}") from exc
    except URLError as exc:
        raise RuntimeError(f"Gemini网络请求失败：{exc}") from exc

    data = json.loads(response_text)
    candidates = data.get("candidates") or []
    if not candidates:
        raise RuntimeError(
            "Gemini响应中没有candidates："
            + json.dumps(data, ensure_ascii=False)[:1500]
        )
    parts = ((candidates[0].get("content") or {}).get("parts") or [])
    content = "".join(
        part.get("text", "")
        for part in parts
        if isinstance(part, dict) and isinstance(part.get("text"), str)
    ).strip()
    if not content:
        raise RuntimeError(
            "Gemini响应中没有文本："
            + json.dumps(data, ensure_ascii=False)[:1500]
        )
    return extract_json(content)


def validate_one_result(raw: dict, source: dict) -> dict:
    record_id = normalize_record_id(raw.get("record_id"))
    if record_id != source["record_id"]:
        raise ValueError(
            f"record_id不一致：期望{source['record_id']}，得到{record_id}"
        )

    code = normalize_text(raw.get("first_level_code"))
    evidence = normalize_text(raw.get("evidence"))
    memo = normalize_text(raw.get("memo"))
    original = source["original_text"]
    is_new_code = parse_bool(raw.get("is_new_code"), default=True)

    if not code:
        raise ValueError("first_level_code为空")
    if code == "文本主旨待细化":
        raise ValueError("禁止使用“文本主旨待细化”作为保底编码")
    if not evidence:
        raise ValueError("evidence为空")

    needs_review = parse_bool(raw.get("needs_review", False))
    review_reason = normalize_text(raw.get("review_reason"))
    if evidence not in original:
        needs_review = True
        reason = "模型给出的证据不是原文逐字片段，需人工复核。"
        review_reason = f"{review_reason}；{reason}".strip("；")

    if len(code) < 6 or len(code) > 80:
        needs_review = True
        reason = "一级编码长度异常，需检查是否过宽或过细。"
        review_reason = f"{review_reason}；{reason}".strip("；")

    shown_reference_codes = {
        item.get("人工一级编码", "")
        for item in source.get("reference_examples", [])
        if isinstance(item, dict)
    }
    if not is_new_code and code not in shown_reference_codes:
        needs_review = True
        reason = "标记为复用旧编码，但编码名称不在本条展示的人工参考案例中。"
        review_reason = f"{review_reason}；{reason}".strip("；")

    return {
        "record_id": source["record_id"],
        "first_level_code": code,
        "is_new_code": is_new_code,
        "evidence": evidence,
        "memo": memo,
        "confidence": clamp_confidence(raw.get("confidence")),
        "needs_review": needs_review,
        "review_reason": review_reason,
        "status": "success",
        "error": "",
    }


def encode_batch_once(
    records: list[dict],
    api_url: str,
    api_key: str,
    model: str,
    temperature: float,
    max_tokens: int,
    timeout: int,
) -> tuple[dict[str, dict], list[str]]:
    response_text = call_gemini_generate_content(
        api_url=api_url,
        api_key=api_key,
        model=model,
        records=records,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout,
    )
    parsed = extract_json(response_text)
    raw_results = parsed.get("results")
    if not isinstance(raw_results, list):
        raise ValueError("JSON顶层缺少results数组")

    source_map = {record["record_id"]: record for record in records}
    valid: dict[str, dict] = {}
    invalid_messages: list[str] = []
    for index, raw in enumerate(raw_results):
        if not isinstance(raw, dict):
            invalid_messages.append(f"results[{index}]不是对象")
            continue
        record_id = normalize_record_id(raw.get("record_id"))
        source = source_map.get(record_id)
        if source is None:
            invalid_messages.append(f"返回了未知record_id={record_id}")
            continue
        if record_id in valid:
            invalid_messages.append(f"record_id={record_id}重复返回")
            continue
        try:
            valid[record_id] = validate_one_result(raw, source)
        except Exception as exc:
            invalid_messages.append(f"record_id={record_id}: {exc}")

    return valid, invalid_messages


def encode_records_with_retries(
    records: list[dict],
    api_url: str,
    api_key: str,
    model: str,
    temperature: float,
    max_tokens: int,
    timeout: int,
    max_retries: int,
) -> dict[str, dict]:
    """先编码整批，再逐条补齐遗漏项；最终失败项明确留空。"""
    completed: dict[str, dict] = {}
    last_error = ""

    for attempt in range(max_retries):
        pending = [r for r in records if r["record_id"] not in completed]
        if not pending:
            break
        try:
            valid, messages = encode_batch_once(
                records=pending,
                api_url=api_url,
                api_key=api_key,
                model=model,
                temperature=temperature,
                max_tokens=max_tokens,
                timeout=timeout,
            )
            completed.update(valid)
            if messages:
                last_error = " | ".join(messages[:20])
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        if len(completed) < len(records):
            retry_sleep(attempt)

    # 批量调用后仍遗漏的记录，逐条再尝试，避免一个敏感文本拖累整批。
    for record in records:
        record_id = record["record_id"]
        if record_id in completed:
            continue
        one_error = last_error
        for attempt in range(max_retries):
            try:
                valid, messages = encode_batch_once(
                    records=[record],
                    api_url=api_url,
                    api_key=api_key,
                    model=model,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    timeout=timeout,
                )
                if record_id in valid:
                    completed[record_id] = valid[record_id]
                    break
                one_error = " | ".join(messages) or "模型遗漏该record_id"
            except Exception as exc:
                one_error = f"{type(exc).__name__}: {exc}"
            retry_sleep(attempt)
        if record_id not in completed:
            completed[record_id] = make_failed_result(record, one_error)

    return completed


# ============================================================
# Stage 2：识别术语 + 豆包联网搜索
# ============================================================

TERM_DETECTION_SYSTEM_PROMPT = """
你是一名计算扎根理论研究助理。现在只判断：一级开放编码是否因网络梗、缩写、专有名词、人物代称、作品名或新兴表达而需要联网检索。

判断原则：
1. 只有当不了解该词会实质影响对用户互动目的的判断时，needs_search 才为 true。
2. 普通人名、常见影视人物名、年份、日常口语，即使属于专名，只要不影响行为目的判断，也不要检索。
3. 术语必须逐字出现在原文中；每条最多3个。
4. 不做二级或三级需求归纳，不修改一级编码。
5. 只输出严格JSON。

输出：
{
  "results": [
    {
      "record_id": "原样返回",
      "interpretation": "贴近原文的语义解释",
      "needs_search": false,
      "still_unclear": false,
      "uncertainty": "无或具体不确定点",
      "suspect_terms": [
        {"term": "原文词语", "reason": "为何可能影响编码", "confidence": 0.0}
      ],
      "memo": "判断说明"
    }
  ]
}
""".strip()


def normalize_suspect_terms(raw_terms: Any, original_text: str, max_terms: int) -> list[dict]:
    terms: list[dict] = []
    seen = set()
    if not isinstance(raw_terms, list):
        return terms
    for item in raw_terms:
        if isinstance(item, str):
            term = normalize_text(item)
            reason = "模型识别为可能需要联网消歧的词语"
        elif isinstance(item, dict):
            term = normalize_text(item.get("term"))
            reason = normalize_text(item.get("reason"))
        else:
            continue
        if not term or term in seen or term not in original_text:
            continue
        seen.add(term)
        confidence = (
            clamp_confidence(item.get("confidence"))
            if isinstance(item, dict) else 0.70
        )
        terms.append({"term": term, "reason": reason, "confidence": confidence})
        if len(terms) >= max(0, max_terms):
            break
    return terms


def detect_terms_batch(
    records: list[dict],
    stage1_map: dict[str, dict],
    api_url: str,
    api_key: str,
    model: str,
    timeout: int,
    max_tokens: int,
    max_terms: int,
    max_retries: int,
) -> dict[str, dict]:
    payload = [
        {
            "record_id": record["record_id"],
            "original_text": record["original_text"],
            "stage1_first_level_code": stage1_map[record["record_id"]].get("first_level_code", ""),
        }
        for record in records
    ]
    user_prompt = (
        "请判断以下文本是否需要联网检索。每个record_id必须且只能返回一次。\n\n"
        + json.dumps(payload, ensure_ascii=False)
    )
    last_error = ""
    for attempt in range(max_retries):
        try:
            parsed = call_gemini_json_prompt(
                api_url=api_url,
                api_key=api_key,
                model=model,
                system_prompt=TERM_DETECTION_SYSTEM_PROMPT,
                user_prompt=user_prompt,
                temperature=1.0,
                max_tokens=max_tokens,
                timeout=timeout,
            )
            raw_results = parsed.get("results")
            if not isinstance(raw_results, list):
                raise ValueError("术语识别结果缺少results数组")
            source_map = {record["record_id"]: record for record in records}
            output: dict[str, dict] = {}
            for raw in raw_results:
                if not isinstance(raw, dict):
                    continue
                record_id = normalize_record_id(raw.get("record_id"))
                source = source_map.get(record_id)
                if source is None or record_id in output:
                    continue
                terms = normalize_suspect_terms(
                    raw.get("suspect_terms"),
                    source["original_text"],
                    max_terms=max_terms,
                )
                needs_search = parse_bool(raw.get("needs_search"), default=bool(terms)) and bool(terms)
                output[record_id] = {
                    "record_id": record_id,
                    "interpretation": normalize_text(raw.get("interpretation")),
                    "needs_search": needs_search,
                    "still_unclear": parse_bool(raw.get("still_unclear")),
                    "uncertainty": normalize_text(raw.get("uncertainty")),
                    "suspect_terms": terms if needs_search else [],
                    "detection_memo": normalize_text(raw.get("memo")),
                    "search_results": [],
                    "status": "terms_detected",
                    "error": "",
                }
            missing = [r["record_id"] for r in records if r["record_id"] not in output]
            if missing:
                raise ValueError(f"术语识别遗漏record_id：{missing[:20]}")
            return output
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt + 1 < max_retries:
                retry_sleep(attempt)

    return {
        record["record_id"]: {
            "record_id": record["record_id"],
            "interpretation": "",
            "needs_search": False,
            "still_unclear": True,
            "uncertainty": "术语识别与语义解释失败",
            "suspect_terms": [],
            "detection_memo": "术语识别失败，未执行自动搜索。",
            "search_results": [],
            "status": "term_detection_failed",
            "error": last_error,
        }
        for record in records
    }


def collect_urls(value: Any, output: list[str]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key.lower() in {"url", "uri", "link"} and isinstance(child, str):
                if child.startswith(("http://", "https://")) and child not in output:
                    output.append(child)
            else:
                collect_urls(child, output)
    elif isinstance(value, list):
        for child in value:
            collect_urls(child, output)


def extract_ark_response_text(data: dict) -> str:
    direct = data.get("output_text")
    if isinstance(direct, str) and direct.strip():
        return direct.strip()
    text_parts: list[str] = []
    for item in data.get("output") or []:
        if not isinstance(item, dict) or item.get("type") != "message":
            continue
        for part in item.get("content") or []:
            if not isinstance(part, dict):
                continue
            text = part.get("text")
            if isinstance(text, str) and text.strip():
                text_parts.append(text.strip())
    return "\n".join(text_parts).strip()


def call_doubao_web_search(
    api_key: str,
    model: str,
    base_url: str,
    term: str,
    original_text: str,
    stage1_code: str,
    max_keyword: int,
    result_limit: int,
    timeout: int,
) -> dict:
    """使用火山方舟Responses API的web_search工具检索并解释术语。"""
    prompt = f"""
你是计算扎根理论研究中的中文网络术语检索助手。
请使用联网搜索核实下列词语在这条微博语境中的含义，并只输出JSON对象。

待检索词：{term}
原文：{original_text}
Stage 1 一级编码：{stage1_code}

输出格式：
{{
  "term": "原词",
  "meaning_cn": "在本条语境下最可能的含义",
  "coding_relevance": "该含义是否以及如何影响一级编码判断",
  "confidence": 0.0,
  "needs_human_review": false,
  "review_reason": ""
}}

要求：不确定时明确写不确定，不得编造来源，不输出JSON之外的内容。
""".strip()
    url = base_url.rstrip("/") + "/responses"
    payload = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt}],
            }
        ],
        "tools": [
            {
                "type": "web_search",
                "max_keyword": max(1, max_keyword),
                "limit": max(1, result_limit),
            }
        ],
        "stream": False,
    }
    request = Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urlopen(
            request,
            timeout=timeout,
            context=get_gemini_ssl_context(),
        ) as response:
            response_text = response.read().decode("utf-8", "replace")
    except HTTPError as exc:
        error_text = exc.read().decode("utf-8", "replace")
        raise RuntimeError(f"豆包搜索HTTP {exc.code}: {error_text[:1800]}") from exc
    except URLError as exc:
        raise RuntimeError(f"豆包搜索网络错误：{exc}") from exc

    data = json.loads(response_text)
    content = extract_ark_response_text(data)
    if not content:
        raise RuntimeError(
            "豆包Responses API没有返回文本："
            + json.dumps(data, ensure_ascii=False)[:1800]
        )
    parsed = extract_json(content)
    urls: list[str] = []
    collect_urls(data, urls)
    return {
        "term": term,
        "meaning_cn": normalize_text(parsed.get("meaning_cn")),
        "coding_relevance": normalize_text(parsed.get("coding_relevance")),
        "confidence": clamp_confidence(parsed.get("confidence")),
        "needs_human_review": parse_bool(parsed.get("needs_human_review")),
        "review_reason": normalize_text(parsed.get("review_reason")),
        "source_urls": urls[:20],
        "status": "search_success",
        "error": "",
    }


def make_search_cache_key(record_id: str, term: str, original_text: str) -> str:
    raw = f"{record_id}\n{term}\n{original_text}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def search_one_with_retries(task: dict, args: argparse.Namespace) -> tuple[str, dict]:
    last_error = ""
    for attempt in range(args.max_retries):
        try:
            result = call_doubao_web_search(
                api_key=args.ark_api_key,
                model=args.ark_model,
                base_url=args.ark_base_url,
                term=task["term"],
                original_text=task["original_text"],
                stage1_code=task["stage1_code"],
                max_keyword=args.doubao_max_keyword,
                result_limit=args.doubao_result_limit,
                timeout=args.timeout,
            )
            return task["cache_key"], result
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt + 1 < args.max_retries:
                retry_sleep(attempt)
    return task["cache_key"], {
        "term": task["term"],
        "meaning_cn": "",
        "coding_relevance": "",
        "confidence": "",
        "needs_human_review": True,
        "review_reason": "豆包联网搜索失败。",
        "source_urls": [],
        "status": "search_failed",
        "error": last_error,
    }


# ============================================================
# Stage 3：依据豆包检索结果复核并回填
# ============================================================

STAGE3_SYSTEM_PROMPT = """
你是一名严格的计算扎根理论研究助理。请使用Stage 2的联网检索结果复核Stage 1一级开放编码。

规则：
1. 仍然只输出一个一级开放编码，不做二级聚类或三级心理需求命名。
2. 只有当检索证据表明Stage 1误解了网络词、专名或语境时才修改；否则原样保留。
3. 最终编码一般15—45个汉字，使用“向AI……”“请求AI……”“与AI分享……”等行为表达。
4. evidence必须逐字来自原文。
5. 搜索失败、来源不足或含义冲突时，不臆测，保留Stage 1并标记人工复核。
6. 只输出严格JSON。

输出：
{
  "results": [
    {
      "record_id": "原样返回",
      "interpretation": "结合检索结果后的语义解释",
      "still_unclear": false,
      "uncertainty": "无或具体不确定点",
      "final_first_level_code": "最终一级编码",
      "changed_from_stage1": false,
      "evidence": "原文证据",
      "memo": "保留或修改的依据",
      "confidence": 0.0,
      "needs_review": false,
      "review_reason": ""
    }
  ]
}
""".strip()


def make_stage3_copy(
    stage1: dict,
    status: str,
    review: bool = False,
    reason: str = "",
    interpretation: str = "",
    still_unclear: bool = False,
    uncertainty: str = "",
) -> dict:
    return {
        "record_id": stage1.get("record_id", ""),
        "interpretation": interpretation,
        "still_unclear": still_unclear,
        "uncertainty": uncertainty,
        "final_first_level_code": stage1.get("first_level_code", ""),
        "changed_from_stage1": False,
        "evidence": stage1.get("evidence", ""),
        "memo": "Stage 2未发现会实质影响编码判断的术语，保留Stage 1编码。",
        "confidence": stage1.get("confidence", ""),
        "needs_review": bool(review or stage1.get("needs_review", False)),
        "review_reason": reason or stage1.get("review_reason", ""),
        "status": status,
        "error": "",
    }


def review_stage3_batch(
    records: list[dict],
    stage1_map: dict[str, dict],
    stage2_map: dict[str, dict],
    api_url: str,
    api_key: str,
    model: str,
    timeout: int,
    max_tokens: int,
    max_retries: int,
) -> dict[str, dict]:
    payload = []
    for record in records:
        record_id = record["record_id"]
        stage1 = stage1_map[record_id]
        stage2 = stage2_map[record_id]
        payload.append(
            {
                "record_id": record_id,
                "original_text": record["original_text"],
                "stage1": {
                    "first_level_code": stage1.get("first_level_code", ""),
                    "evidence": stage1.get("evidence", ""),
                    "memo": stage1.get("memo", ""),
                },
                "suspect_terms": stage2.get("suspect_terms", []),
                "doubao_web_search_results": stage2.get("search_results", []),
            }
        )
    user_prompt = (
        "请复核以下记录。每个record_id必须且只能返回一次。\n\n"
        + json.dumps(payload, ensure_ascii=False)
    )
    last_error = ""
    for attempt in range(max_retries):
        try:
            parsed = call_gemini_json_prompt(
                api_url=api_url,
                api_key=api_key,
                model=model,
                system_prompt=STAGE3_SYSTEM_PROMPT,
                user_prompt=user_prompt,
                temperature=1.0,
                max_tokens=max_tokens,
                timeout=timeout,
            )
            raw_results = parsed.get("results")
            if not isinstance(raw_results, list):
                raise ValueError("Stage 3结果缺少results数组")
            source_map = {record["record_id"]: record for record in records}
            output: dict[str, dict] = {}
            for raw in raw_results:
                if not isinstance(raw, dict):
                    continue
                record_id = normalize_record_id(raw.get("record_id"))
                source = source_map.get(record_id)
                if source is None or record_id in output:
                    continue
                stage1 = stage1_map[record_id]
                code = normalize_text(raw.get("final_first_level_code"))
                evidence = normalize_text(raw.get("evidence"))
                if not code or not evidence:
                    continue
                needs_review = parse_bool(raw.get("needs_review"))
                reason = normalize_text(raw.get("review_reason"))
                if evidence not in source["original_text"]:
                    needs_review = True
                    reason = (reason + "；Stage 3证据不是原文逐字片段").strip("；")
                output[record_id] = {
                    "record_id": record_id,
                    "interpretation": normalize_text(raw.get("interpretation")),
                    "still_unclear": parse_bool(raw.get("still_unclear")),
                    "uncertainty": normalize_text(raw.get("uncertainty")),
                    "final_first_level_code": code,
                    "changed_from_stage1": code != stage1.get("first_level_code", ""),
                    "evidence": evidence,
                    "memo": normalize_text(raw.get("memo")),
                    "confidence": clamp_confidence(raw.get("confidence")),
                    "needs_review": needs_review,
                    "review_reason": reason,
                    "status": "stage3_success",
                    "error": "",
                }
            missing = [r["record_id"] for r in records if r["record_id"] not in output]
            if missing:
                raise ValueError(f"Stage 3遗漏record_id：{missing[:20]}")
            return output
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt + 1 < max_retries:
                retry_sleep(attempt)

    return {
        record["record_id"]: {
            **make_stage3_copy(
                stage1_map[record["record_id"]],
                status="stage3_failed",
                review=True,
                reason="Stage 3 API复核失败，暂时保留Stage 1编码。",
                interpretation=stage2_map[record["record_id"]].get("interpretation", ""),
                still_unclear=True,
                uncertainty="Stage 3复核失败",
            ),
            "error": last_error,
        }
        for record in records
    }


def export_full_pipeline_excel(
    source_df: pd.DataFrame,
    stage1_map: dict[str, dict],
    stage2_map: dict[str, dict],
    stage3_map: dict[str, dict],
    output_path: Path,
) -> None:
    rows = []
    for _, row in source_df.iterrows():
        record_id = row["_record_id"]
        s1 = stage1_map.get(record_id, {})
        s2 = stage2_map.get(record_id, {})
        s3 = stage3_map.get(record_id, {})
        search_results = s2.get("search_results", [])
        source_urls = []
        for item in search_results if isinstance(search_results, list) else []:
            if isinstance(item, dict):
                for url in item.get("source_urls", []) or []:
                    if url not in source_urls:
                        source_urls.append(url)
        rows.append(
            {
                SEQ_COLUMN: row[SEQ_COLUMN],
                ID_COLUMN: row[ID_COLUMN],
                TEXT_COLUMN: row[TEXT_COLUMN] if not pd.isna(row[TEXT_COLUMN]) else "",
                "Stage1一级编码": s1.get("first_level_code", ""),
                "Stage1编码证据": s1.get("evidence", ""),
                "Stage1编码备忘录": s1.get("memo", ""),
                "Stage1状态": s1.get("status", "pending"),
                "Stage2是否需搜索": s2.get("needs_search", ""),
                "Stage2疑似术语": json.dumps(s2.get("suspect_terms", []), ensure_ascii=False),
                "Stage2豆包搜索结果": json.dumps(search_results, ensure_ascii=False),
                "Stage2来源链接": "\n".join(source_urls),
                "Stage2状态": s2.get("status", "pending"),
                "Stage2错误": s2.get("error", ""),
                "Stage3最终一级编码": s3.get("final_first_level_code", ""),
                "Stage3是否修改Stage1": s3.get("changed_from_stage1", ""),
                "Stage3证据": s3.get("evidence", ""),
                "Stage3复核备忘录": s3.get("memo", ""),
                "Stage3置信度": s3.get("confidence", ""),
                "Stage3是否需人工复核": s3.get("needs_review", ""),
                "Stage3复核原因": s3.get("review_reason", ""),
                "Stage3状态": s3.get("status", "pending"),
                "Stage3错误": s3.get("error", ""),
                "最终采用一级编码": s3.get("final_first_level_code", s1.get("first_level_code", "")),
            }
        )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.stem + "_tmp.xlsx")
    pd.DataFrame(rows).to_excel(temp_path, index=False)
    os.replace(temp_path, output_path)


def attach_content_ids_from_references(
    records: list[dict],
    reference_paths: list[Path],
    sheet_name: str | int = 0,
) -> None:
    """从既有Stage文件按原文回查content_id；重复文本按最近物理行消歧。"""
    import gc
    import xml.etree.ElementTree as ET
    import zipfile

    def clear_xml_element(element: Any) -> None:
        """释放已解析XML节点。"""
        element.clear()

    existing_paths = [path for path in reference_paths if path and path.exists()]
    if not existing_paths:
        for record in records:
            record["content_id"] = None
            record["content_id_lookup_status"] = (
                "blank_text_unresolved"
                if not record["original_text"]
                else "reference_files_missing"
            )
            record["content_id_reference_file"] = ""
            record["matched_reference_excel_row"] = None
        print(
            "⚠️ 未找到Stage 1/2/3参考表：content_id暂时留空，"
            "但不中断Stage 2豆包搜索与Stage 3。\n"
            "若需补齐content_id，后续将stage1.xlsx、stage2.xlsx、"
            "stage3.xlsx放到脚本同目录后重新运行即可。",
            flush=True,
        )
        return

    records_by_text: dict[str, list[dict]] = {}
    for record in records:
        if record["original_text"]:
            records_by_text.setdefault(record["original_text"], []).append(record)
        else:
            record["content_id"] = None
            record["content_id_lookup_status"] = "blank_text_unresolved"

    candidate_map: dict[str, list[tuple[int, Any, str]]] = {
        text: [] for text in records_by_text
    }

    # 大表采用XLSX内部XML流式解析，避免一次性加载37万行和全部共享字符串。
    for reference_path in existing_paths:
        active_texts = {
            text
            for text, matching_records in records_by_text.items()
            if len({item[1] for item in candidate_map[text]}) < len(matching_records)
        }
        if not active_texts:
            break
        with zipfile.ZipFile(reference_path) as archive:
            names = set(archive.namelist())
            target_shared_strings: dict[int, str] = {}
            if "xl/sharedStrings.xml" in names:
                shared_index = -1
                with archive.open("xl/sharedStrings.xml") as shared_file:
                    context = ET.iterparse(shared_file, events=("start", "end"))
                    _, shared_root = next(context)
                    for event, element in context:
                        if event == "end" and (element.tag.endswith("}si") or element.tag == "si"):
                            shared_index += 1
                            text = "".join(
                                child.text or ""
                                for child in element.iter()
                                if child.tag.endswith("}t") or child.tag == "t"
                            )
                            normalized = normalize_text(text)
                            if normalized in active_texts:
                                target_shared_strings[shared_index] = normalized
                            clear_xml_element(element)
                            if shared_index % 1000 == 0:
                                shared_root.clear()

            worksheet_xml = (
                f"xl/worksheets/sheet{sheet_name + 1}.xml"
                if isinstance(sheet_name, int)
                else "xl/worksheets/sheet1.xml"
            )
            if worksheet_xml not in names:
                continue
            with archive.open(worksheet_xml) as worksheet_file:
                context = ET.iterparse(worksheet_file, events=("start", "end"))
                _, worksheet_root = next(context)
                for event, row_element in context:
                    if event != "end" or not (row_element.tag.endswith("}row") or row_element.tag == "row"):
                        continue
                    excel_row = int(row_element.attrib.get("r", "0") or 0)
                    content_id = None
                    reference_text = ""
                    for cell in row_element:
                        if not (cell.tag.endswith("}c") or cell.tag == "c"):
                            continue
                        cell_ref = cell.attrib.get("r", "")
                        column = re.sub(r"\d+", "", cell_ref)
                        if column not in {"A", "B"}:
                            continue
                        cell_type = cell.attrib.get("t", "")
                        value_text = next((
                            child.text
                            for child in cell.iter()
                            if child.tag.endswith("}v") or child.tag == "v"
                        ), None)
                        if column == "A" and value_text is not None:
                            try:
                                number = float(value_text)
                                content_id = int(number) if number.is_integer() else number
                            except Exception:
                                content_id = value_text
                        elif column == "B":
                            if cell_type == "s" and value_text is not None:
                                try:
                                    reference_text = target_shared_strings.get(int(value_text), "")
                                except Exception:
                                    reference_text = ""
                            elif cell_type == "inlineStr":
                                reference_text = normalize_text("".join(
                                    child.text or ""
                                    for child in cell.iter()
                                    if child.tag.endswith("}t") or child.tag == "t"
                                ))
                            else:
                                reference_text = normalize_text(value_text)
                    if content_id is not None and reference_text in candidate_map:
                        candidate = (excel_row, content_id, reference_path.name)
                        if candidate not in candidate_map[reference_text]:
                            candidate_map[reference_text].append(candidate)
                    clear_xml_element(row_element)
                    if excel_row % 1000 == 0:
                        worksheet_root.clear()
        gc.collect()

    missing = []
    for text, matching_records in records_by_text.items():
        available = list(candidate_map.get(text, []))
        matching_records.sort(key=lambda r: int(float(r["source_excel_row"])))
        used_content_ids = set()
        for record in matching_records:
            candidates = [item for item in available if item[1] not in used_content_ids]
            if not candidates:
                missing.append(record["source_excel_row"])
                record["content_id"] = None
                record["content_id_lookup_status"] = "text_not_found"
                continue
            source_row = int(float(record["source_excel_row"]))
            matched_row, content_id, source_file = min(
                candidates,
                key=lambda item: abs(item[0] - source_row),
            )
            used_content_ids.add(content_id)
            record["content_id"] = content_id
            record["content_id_lookup_status"] = "exact_text_nearest_row"
            record["content_id_reference_file"] = source_file
            record["matched_reference_excel_row"] = matched_row

    if missing:
        print(
            f"⚠️ content_id回查：仍有{len(missing)}条非空原文未在Stage参考文件中找到；"
            "这些记录的content_id将留空，并在JSON中保留原Excel行号供人工回查。\n"
            f"前30个原Excel行号：{missing[:30]}",
            flush=True,
        )


STAGE_CODE_COLUMNS = [
    "content_id", "original_text", "interpretation", "needs_search",
    "still_unclear", "uncertainty", "suspect_terms",
    "code_1", "evidence_1", "memo_1", "confidence_1", "is_new_1",
    "code_2", "evidence_2", "memo_2", "confidence_2", "is_new_2",
    "code_3", "evidence_3", "memo_3", "confidence_3", "is_new_3",
]

STAGE2_COLUMNS = [
    "content_id", "original_text", "interpretation", "needs_search",
    "suspect_terms", "doubao_notes_n", "doubao_notes",
    "term_1", "meaning_1", "confidence_1",
    "term_2", "meaning_2", "confidence_2",
    "term_3", "meaning_3", "confidence_3",
]

STAGE3_COLUMNS = [
    "content_id", "original_text", "interpretation", "needs_search",
    "still_unclear", "uncertainty", "terms_json",
    "code_1", "evidence_1", "memo_1", "confidence_1", "is_new_1",
    "code_2", "evidence_2", "memo_2", "confidence_2", "is_new_2",
    "code_3", "evidence_3", "memo_3", "confidence_3", "is_new_3",
]


def make_code_slots(code: str, evidence: str, memo: str, confidence: Any, is_new: Any) -> dict:
    output = {}
    for index in range(1, 4):
        output[f"code_{index}"] = code if index == 1 else None
        output[f"evidence_{index}"] = evidence if index == 1 else None
        output[f"memo_{index}"] = memo if index == 1 else None
        output[f"confidence_{index}"] = confidence if index == 1 else None
        output[f"is_new_{index}"] = is_new if index == 1 else None
    return output


def build_template_rows(
    records: list[dict],
    stage1_map: dict[str, dict],
    stage2_map: dict[str, dict],
    stage3_map: dict[str, dict],
) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
    stage1_rows, stage2_rows, stage3_rows, json_rows = [], [], [], []
    for record in records:
        record_id = record["record_id"]
        content_id = record.get("content_id")
        original_text = record["original_text"]
        s1 = stage1_map.get(record_id, {})
        s2 = stage2_map.get(record_id, {})
        s3 = stage3_map.get(record_id, {})
        interpretation = s2.get("interpretation") or s1.get("memo", "")
        suspect_terms = s2.get("suspect_terms", []) or []

        stage1_row = {
            "content_id": content_id,
            "original_text": original_text,
            "interpretation": interpretation,
            "needs_search": s2.get("needs_search", False),
            "still_unclear": s2.get("still_unclear", False),
            "uncertainty": s2.get("uncertainty") or None,
            "suspect_terms": json.dumps(suspect_terms, ensure_ascii=False),
            **make_code_slots(
                s1.get("first_level_code", ""),
                s1.get("evidence", ""),
                s1.get("memo", ""),
                s1.get("confidence", ""),
                s1.get("is_new_code", ""),
            ),
        }
        stage1_rows.append(stage1_row)

        search_results = s2.get("search_results", []) or []
        doubao_notes = [
            {
                "term": item.get("term", ""),
                "meaning": item.get("meaning_cn", ""),
                "confidence": item.get("confidence", ""),
            }
            for item in search_results
            if isinstance(item, dict) and item.get("status") == "search_success"
        ]
        unresolved = bool(
            s2.get("needs_search") and s2.get("status") != "search_complete"
        )
        stage2_row = {
            "content_id": content_id,
            "original_text": original_text,
            "interpretation": interpretation,
            "needs_search": unresolved,
            "suspect_terms": json.dumps(suspect_terms if unresolved else [], ensure_ascii=False),
            "doubao_notes_n": len(doubao_notes),
            "doubao_notes": json.dumps(doubao_notes, ensure_ascii=False),
        }
        for index in range(1, 4):
            item = doubao_notes[index - 1] if index <= len(doubao_notes) else {}
            stage2_row[f"term_{index}"] = item.get("term")
            stage2_row[f"meaning_{index}"] = item.get("meaning")
            stage2_row[f"confidence_{index}"] = item.get("confidence")
        stage2_rows.append(stage2_row)

        terms_json = [
            {
                "term": item.get("term", ""),
                "meaning_cn": item.get("meaning_cn", ""),
                "source": "doubao",
                "confidence": item.get("confidence", ""),
            }
            for item in search_results
            if isinstance(item, dict) and item.get("status") == "search_success"
        ]
        final_code = s3.get("final_first_level_code", s1.get("first_level_code", ""))
        changed = parse_bool(s3.get("changed_from_stage1"))
        stage3_interpretation = s3.get("interpretation") or interpretation
        stage3_row = {
            "content_id": content_id,
            "original_text": original_text,
            "interpretation": stage3_interpretation,
            "needs_search": s2.get("needs_search", False),
            "still_unclear": s3.get("still_unclear", s2.get("still_unclear", False)),
            "uncertainty": s3.get("uncertainty") or s2.get("uncertainty") or None,
            "terms_json": json.dumps(terms_json, ensure_ascii=False),
            **make_code_slots(
                final_code,
                s3.get("evidence", s1.get("evidence", "")),
                s3.get("memo", s1.get("memo", "")),
                s3.get("confidence", s1.get("confidence", "")),
                True if changed else s1.get("is_new_code", ""),
            ),
        }
        stage3_rows.append(stage3_row)

        json_rows.append(
            {
                "content_id": content_id,
                "source_excel_row": record["source_excel_row"],
                "content_id_lookup_status": record.get("content_id_lookup_status", ""),
                "content_id_reference_file": record.get("content_id_reference_file", ""),
                "matched_reference_excel_row": record.get("matched_reference_excel_row"),
                "original_text": original_text,
                "stage1": stage1_row,
                "stage2": stage2_row,
                "stage2_search_details": search_results,
                "stage3": stage3_row,
            }
        )
    return stage1_rows, stage2_rows, stage3_rows, json_rows


def export_stage_template_files(
    records: list[dict],
    stage1_map: dict[str, dict],
    stage2_map: dict[str, dict],
    stage3_map: dict[str, dict],
    output_dir: Path,
) -> tuple[Path, Path, Path, Path]:
    stage1_rows, stage2_rows, stage3_rows, json_rows = build_template_rows(
        records, stage1_map, stage2_map, stage3_map
    )
    stage1_path = output_dir / "stage1.xlsx"
    stage2_path = output_dir / "stage2.xlsx"
    stage3_path = output_dir / "stage3.xlsx"
    json_path = output_dir / "stage123_results.json"
    outputs = [
        (stage1_path, pd.DataFrame(stage1_rows, columns=STAGE_CODE_COLUMNS)),
        (stage2_path, pd.DataFrame(stage2_rows, columns=STAGE2_COLUMNS)),
        (stage3_path, pd.DataFrame(stage3_rows, columns=STAGE3_COLUMNS)),
    ]
    for path, dataframe in outputs:
        temp_path = path.with_name(path.stem + "_tmp.xlsx")
        dataframe.to_excel(temp_path, index=False)
        os.replace(temp_path, path)
    atomic_write_json(
        json_path,
        {
            "metadata": {
                "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "records": len(json_rows),
                "pipeline": "Stage1 open coding -> Doubao web search -> Stage3 review",
            },
            "results": json_rows,
        },
    )
    return stage1_path, stage2_path, stage3_path, json_path


def export_stage1_template_file(
    records: list[dict],
    stage1_map: dict[str, dict],
    stage2_map: dict[str, dict],
    output_dir: Path,
) -> Path:
    """Stage 1待搜索对象识别完成后立即导出，不等豆包搜索。"""
    stage1_rows, _, _, _ = build_template_rows(
        records, stage1_map, stage2_map, {}
    )
    output_path = output_dir / "stage1.xlsx"
    temp_path = output_path.with_name("stage1_tmp.xlsx")
    pd.DataFrame(stage1_rows, columns=STAGE_CODE_COLUMNS).to_excel(
        temp_path, index=False
    )
    os.replace(temp_path, output_path)
    return output_path


# ============================================================
# 主流程
# ============================================================

def split_batches(records: list[dict], batch_size: int) -> list[list[dict]]:
    return [records[i:i + batch_size] for i in range(0, len(records), batch_size)]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="计算扎根理论 Stage 1→豆包搜索→Stage 3 完整流程")
    parser.add_argument("--input", default=DEFAULT_INPUT_XLSX, help="待编码.xlsx路径")
    parser.add_argument("--coding-book", default=DEFAULT_CODING_BOOK_XLSX, help="人工开放编码参考手册路径")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="输出目录")
    parser.add_argument("--sheet", default="0", help="工作表序号或名称，默认0")
    parser.add_argument("--coding-book-sheet", default="0", help="编码手册工作表序号或名称")
    parser.add_argument("--reference-top-k", type=int, default=3, help="每条文本展示的相似人工案例数")
    parser.add_argument("--reference-min-score", type=float, default=0.16, help="人工案例最低文本相似度")
    parser.add_argument("--api-url", default=DEFAULT_API_URL)
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument("--batch-size", type=int, default=10)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--max-retries", type=int, default=3)
    parser.add_argument("--max-tokens", type=int, default=6000)
    parser.add_argument("--temperature", type=float, default=1.0)
    parser.add_argument("--timeout", type=int, default=180)
    parser.add_argument("--limit", type=int, default=0, help="只处理前N条待处理记录，用于测试")
    parser.add_argument("--rerun-failed", action="store_true", help="重新请求checkpoint中的失败项")
    parser.add_argument("--skip-api-test", action="store_true", help="跳过启动前的Gemini密钥与模型连接测试")
    parser.add_argument("--stage1-only", action="store_true", help="只运行Stage 1，不执行豆包搜索和Stage 3")
    parser.add_argument("--stage2-batch-size", type=int, default=20, help="Stage 2术语识别批大小")
    parser.add_argument("--stage3-batch-size", type=int, default=10, help="Stage 3复核批大小")
    parser.add_argument("--max-search-terms", type=int, default=3, help="每条原文最多检索术语数")
    parser.add_argument("--ark-base-url", default=os.getenv("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3"))
    parser.add_argument("--ark-model", default=os.getenv("ARK_MODEL", ""), help="支持Responses API和web_search的豆包模型或推理接入点")
    parser.add_argument("--doubao-workers", type=int, default=2)
    parser.add_argument("--doubao-max-keyword", type=int, default=2)
    parser.add_argument("--doubao-result-limit", type=int, default=5)
    parser.add_argument("--rerun-stage2-failed", action="store_true")
    parser.add_argument("--rerun-stage3-failed", action="store_true")
    parser.add_argument(
        "--rerun-restored",
        action="store_true",
        help="一次性强制重跑53条已恢复原文记录的Stage 1、豆包搜索和Stage 3",
    )
    parser.add_argument("--dry-run", action="store_true", help="只检查文件和参数，不调用API")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    coding_book_path = Path(args.coding_book).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"找不到输入文件：{input_path}")
    if not coding_book_path.exists():
        raise FileNotFoundError(f"找不到人工编码手册：{coding_book_path}")

    sheet: str | int
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    source_df, records = load_pending_records(input_path, sheet=sheet)
    reject_unrecovered_placeholders(records)

    book_sheet: str | int
    book_sheet = int(args.coding_book_sheet) if str(args.coding_book_sheet).isdigit() else args.coding_book_sheet
    coding_book, coding_book_codes = load_coding_book(coding_book_path, sheet=book_sheet)
    for record in records:
        record["reference_examples"] = retrieve_reference_examples(
            record["original_text"],
            coding_book,
            top_k=max(0, args.reference_top_k),
            min_score=max(0.0, args.reference_min_score),
        )

    model_tag = safe_filename(args.model)
    checkpoint_path = output_dir / f"{model_tag}_一级编码_checkpoint.json"
    output_xlsx = output_dir / "待编码_一级编码结果.xlsx"

    print(f"输入文件：{input_path}")
    print(f"人工编码手册：{coding_book_path}")
    print(f"有效人工示例：{len(coding_book)}条；既有一级编码：{len(coding_book_codes)}个")
    print(f"总记录数：{len(records)}")
    print(f"原文为空：{sum(not r['original_text'] for r in records)}")
    print(f"输出文件：{output_xlsx}")
    print(f"断点文件：{checkpoint_path}")

    if args.dry_run:
        print("\n前3条文本检索到的人工参考案例：")
        for record in records[:3]:
            print(json.dumps({
                "record_id": record["record_id"],
                "original_text": record["original_text"],
                "reference_examples": record["reference_examples"],
            }, ensure_ascii=False, indent=2))
        print("Dry run完成：输入表、编码手册及案例检索均正常，未调用API。")
        return

    api_key = os.getenv("GEMINI_API_KEY", "")
    if not api_key:
        raise RuntimeError(
            "没有检测到API Key。请在脚本同目录的.env中设置：\n"
            "GEMINI_API_KEY=你的API_KEY"
        )

    if not args.skip_api_test:
        print(f"[连接测试] 正在验证Google Gemini模型：{args.model}", flush=True)
        model_info = test_gemini_connection(
            api_key=api_key,
            model=args.model,
            timeout=min(max(5, args.timeout), 20),
        )
        print(
            f"[连接测试通过] {model_info.get('displayName', args.model)}；"
            "密钥、模型与接口均可用。",
            flush=True,
        )

    result_map = load_checkpoint(checkpoint_path)
    invalidate_changed_original_text(result_map, records, "Stage 1")
    if args.rerun_restored:
        force_remove_record_ids(
            result_map,
            RESTORED_RECORD_IDS,
            "Stage 1强制重跑恢复记录",
        )

    # 空原文直接记录，不调用API。
    for record in records:
        if not record["original_text"]:
            result_map[record["record_id"]] = make_blank_result(record)

    success_statuses = {"success", "blank_text"}
    pending = []
    for record in records:
        previous = result_map.get(record["record_id"], {})
        status = previous.get("status")
        if status in success_statuses:
            continue
        if status == "api_failed" and not args.rerun_failed:
            continue
        if record["original_text"]:
            pending.append(record)

    if args.limit > 0:
        pending = pending[:args.limit]

    print(f"已从checkpoint恢复：{len(result_map)}条")
    print(f"本次准备请求：{len(pending)}条")

    if not pending:
        atomic_write_json(
            checkpoint_path,
            {
                "input_file": str(input_path),
                "model": args.model,
                "api_url_hash": hashlib.sha256(
                    args.api_url.encode("utf-8")
                ).hexdigest()[:12],
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "results": result_map,
            },
        )
        export_excel(source_df, result_map, output_xlsx)
        print("没有新的待处理记录，已同步checkpoint并重新导出结果Excel。")
        return

    batches = split_batches(pending, max(1, args.batch_size))
    started = time.perf_counter()
    finished_batches = 0
    finished_records = 0
    total_batches = len(batches)
    total_pending_records = len(pending)

    print(
        f"开始一级开放编码：共{total_pending_records}条，"
        f"分为{total_batches}批；batch_size={max(1, args.batch_size)}，"
        f"workers={max(1, args.workers)}",
        flush=True,
    )
    print("API请求已发出；若模型响应较慢，每10秒会打印一次等待状态。", flush=True)

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {
            executor.submit(
                encode_records_with_retries,
                batch,
                args.api_url,
                api_key,
                args.model,
                args.temperature,
                args.max_tokens,
                args.timeout,
                args.max_retries,
            ): batch
            for batch in batches
        }

        unfinished_futures = set(future_map)
        while unfinished_futures:
            done_futures, unfinished_futures = wait(
                unfinished_futures,
                timeout=10,
                return_when=FIRST_COMPLETED,
            )

            # 即使一批尚未返回，也持续告诉用户程序仍在等待API。
            if not done_futures:
                elapsed = time.perf_counter() - started
                active_requests = min(max(1, args.workers), len(unfinished_futures))
                print(
                    f"[等待API] 已运行{elapsed:.0f}秒；"
                    f"已完成{finished_batches}/{total_batches}批（{finished_records}/{total_pending_records}条）；"
                    f"当前最多{active_requests}个请求正在执行。",
                    flush=True,
                )
                continue

            for future in done_futures:
                batch = future_map[future]
                try:
                    batch_results = future.result()
                except Exception as exc:
                    batch_results = {
                        record["record_id"]: make_failed_result(record, str(exc))
                        for record in batch
                    }

                result_map.update(batch_results)
                atomic_write_json(
                    checkpoint_path,
                    {
                        "input_file": str(input_path),
                        "model": args.model,
                        "api_url_hash": hashlib.sha256(args.api_url.encode("utf-8")).hexdigest()[:12],
                        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "results": result_map,
                    },
                )

                finished_batches += 1
                finished_records += len(batch)
                elapsed = time.perf_counter() - started
                percent = 100.0 * finished_records / total_pending_records
                average_batch_seconds = elapsed / finished_batches
                eta_seconds = average_batch_seconds * (total_batches - finished_batches)
                success_count = sum(r.get("status") == "success" for r in result_map.values())
                failed_count = sum(r.get("status") == "api_failed" for r in result_map.values())
                print(
                    f"[进度 {percent:5.1f}%] {finished_batches}/{total_batches}批，"
                    f"{finished_records}/{total_pending_records}条；"
                    f"累计成功={success_count}，失败={failed_count}；"
                    f"已用时={elapsed / 60:.1f}分钟，预计剩余={eta_seconds / 60:.1f}分钟",
                    flush=True,
                )

                # 每5批导出一次，防止程序中断时只留下JSON。
                if finished_batches % 5 == 0:
                    export_excel(source_df, result_map, output_xlsx)
                    print(f"[自动保存] 已更新结果Excel：{output_xlsx}", flush=True)

    export_excel(source_df, result_map, output_xlsx)
    elapsed = time.perf_counter() - started

    success_count = sum(r.get("status") == "success" for r in result_map.values())
    failed_count = sum(r.get("status") == "api_failed" for r in result_map.values())
    blank_count = sum(r.get("status") == "blank_text" for r in result_map.values())
    pending_count = len(records) - success_count - failed_count - blank_count

    print("\n一级开放编码完成")
    print(f"成功：{success_count}")
    print(f"API失败：{failed_count}")
    print(f"原文为空：{blank_count}")
    print(f"尚未处理：{pending_count}")
    print(f"耗时：{elapsed / 60:.1f}分钟")
    print(f"结果Excel：{output_xlsx}")
    print(f"checkpoint：{checkpoint_path}")
    if failed_count:
        print("如需重跑失败项：python 一级开放编码_待编码.py --rerun-failed")


def main_full_pipeline() -> None:
    """先复用原Stage 1流程，再从其checkpoint继续Stage 2和Stage 3。"""
    args = parse_args()
    main()
    if args.dry_run or args.stage1_only:
        return

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    sheet: str | int = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    source_df, records = load_pending_records(input_path, sheet=sheet)
    reject_unrecovered_placeholders(records)
    # 用户后续会自行与主数据匹配；因此这里直接使用输入表中的
    # “原Excel行号”作为临时content_id，不再读取任何旧Stage结果。
    for record in records:
        source_id = normalize_record_id(record["source_excel_row"])
        record["content_id"] = int(source_id) if source_id.isdigit() else source_id
        record["content_id_lookup_status"] = "source_excel_row_as_temporary_id"
        record["content_id_reference_file"] = ""
        record["matched_reference_excel_row"] = None
    print("临时content_id：直接使用待编码表的‘原Excel行号’。", flush=True)
    record_map = {record["record_id"]: record for record in records}

    model_tag = safe_filename(args.model)
    stage1_checkpoint = output_dir / f"{model_tag}_一级编码_checkpoint.json"
    stage2_checkpoint = output_dir / f"{model_tag}_stage2_术语与豆包搜索_checkpoint.json"
    search_cache_path = output_dir / "doubao_web_search_cache.json"
    stage3_checkpoint = output_dir / f"{model_tag}_stage3_复核回填_checkpoint.json"

    stage1_map = load_checkpoint(stage1_checkpoint)
    invalidate_changed_original_text(stage1_map, records, "Stage 1")
    eligible_records = [
        record
        for record in records
        if stage1_map.get(record["record_id"], {}).get("status") == "success"
    ]
    # --limit 同时限制Stage 2/3，避免Stage 1已有全量checkpoint时，
    # “测试3条”却意外触发全量豆包搜索。
    if args.limit > 0:
        eligible_records = eligible_records[:args.limit]
    if not eligible_records:
        print("没有Stage 1成功记录，暂不运行Stage 2与Stage 3。", flush=True)
        return

    gemini_key = os.getenv("GEMINI_API_KEY", "")
    print(
        f"\n========== Stage 1B：识别待搜索对象 =========="
        f"\n可进入待搜索识别的记录：{len(eligible_records)}条",
        flush=True,
    )

    # ---------- Stage 1B：识别真正需要联网消歧的词 ----------
    stage2_map = load_checkpoint(stage2_checkpoint)
    search_cache = load_checkpoint(search_cache_path)
    stage3_map = load_checkpoint(stage3_checkpoint)
    invalidate_changed_original_text(stage2_map, records, "Stage 2")
    invalidate_changed_original_text(stage3_map, records, "Stage 3")
    if args.rerun_restored:
        force_remove_record_ids(
            stage2_map,
            RESTORED_RECORD_IDS,
            "Stage 1B/Stage 2强制重跑恢复记录",
        )
        force_remove_record_ids(
            stage3_map,
            RESTORED_RECORD_IDS,
            "Stage 3强制重跑恢复记录",
        )
        print(
            "[豆包缓存] 恢复记录本次将绕过旧缓存，"
            "需要搜索的术语会重新联网请求。",
            flush=True,
        )

    # 流水线启动前就检查豆包配置，避免等Stage 1B全部完成后才报错。
    ark_key = os.getenv("ARK_API_KEY", "")
    if not ark_key:
        raise RuntimeError(
            "未检测到ARK_API_KEY。为避免Stage 1B跑完后无法自动搜索，"
            "流程已在启动前停止。\n请在.env中配置ARK_API_KEY后重新运行。"
        )
    if not args.ark_model:
        raise RuntimeError(
            "未检测到ARK_MODEL。请在.env中配置支持Responses API和"
            "web_search的豆包模型或推理接入点。"
        )
    args.ark_api_key = ark_key
    print(
        f"[流水线就绪] Gemini待搜索识别 → 豆包联网搜索；"
        f"每批识别完立即搜索。",
        flush=True,
    )

    def save_stage2_state() -> None:
        atomic_write_json(
            stage2_checkpoint,
            {
                "stage": "stage1_detection_and_stage2_doubao_search",
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "results": stage2_map,
            },
        )

    def save_search_cache() -> None:
        atomic_write_json(
            search_cache_path,
            {
                "stage": "doubao_web_search_cache",
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "results": search_cache,
            },
        )

    def save_stage3_state() -> None:
        atomic_write_json(
            stage3_checkpoint,
            {
                "stage": "stage3_search_informed_review",
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "results": stage3_map,
            },
        )

    def review_searched_batch(batch: list[dict], label: str) -> None:
        """本批一旦搜索完成，立即将搜索证据送回Gemini复核。"""
        candidates = [
            record for record in batch
            if stage2_map.get(record["record_id"], {}).get("status")
            == "search_complete"
            and stage3_map.get(record["record_id"], {}).get("status")
            != "stage3_success"
        ]
        if not candidates:
            print(
                f"[Stage 3批次 {label}] 本批无新增搜索或已复核，"
                "无需再请求Gemini。",
                flush=True,
            )
            return

        review_batches = split_batches(
            candidates, max(1, args.stage3_batch_size)
        )
        print(
            f"[Stage 3批次启动 {label}] 有{len(candidates)}条搜索成功记录"
            f"需要回去复核，共{len(review_batches)}批。",
            flush=True,
        )
        review_started = time.perf_counter()
        for review_index, review_batch in enumerate(review_batches, start=1):
            batch_started = time.perf_counter()
            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(
                    review_stage3_batch,
                    records=review_batch,
                    stage1_map=stage1_map,
                    stage2_map=stage2_map,
                    api_url=args.api_url,
                    api_key=gemini_key,
                    model=args.model,
                    timeout=args.timeout,
                    max_tokens=args.max_tokens,
                    max_retries=args.max_retries,
                )
                while True:
                    done, _ = wait(
                        {future}, timeout=10, return_when=FIRST_COMPLETED
                    )
                    if done:
                        reviewed = future.result()
                        break
                    print(
                        f"[等待Stage 3 {label} "
                        f"{review_index}/{len(review_batches)}] "
                        f"本批已运行{time.perf_counter()-batch_started:.0f}秒；"
                        f"正在复核{len(review_batch)}条搜索记录。",
                        flush=True,
                    )
            stage3_map.update(reviewed)
            save_stage3_state()
            completed = sum(
                len(item) for item in review_batches[:review_index]
            )
            elapsed = time.perf_counter() - review_started
            percent = 100.0 * completed / len(candidates)
            eta = elapsed / completed * (len(candidates) - completed)
            print(
                f"[Stage 3进度 {label} {percent:5.1f}%] "
                f"{completed}/{len(candidates)}条；"
                f"已用时={elapsed/60:.1f}分钟；"
                f"预计剩余={eta/60:.1f}分钟。",
                flush=True,
            )

    def search_detected_batch(batch: list[dict], label: str) -> None:
        """一批Stage 1B识别完成后立即搜索，并将结果回填Stage 2 checkpoint。"""
        tasks: list[dict] = []
        task_keys_by_record: dict[str, list[str]] = {}
        records_needing_search = 0

        for record in batch:
            record_id = record["record_id"]
            stage2 = stage2_map.get(record_id, {})
            if not stage2.get("needs_search"):
                stage2["status"] = (
                    "term_detection_failed"
                    if stage2.get("status") == "term_detection_failed"
                    else "no_search_needed"
                )
                stage2["search_results"] = []
                stage2_map[record_id] = stage2
                task_keys_by_record[record_id] = []
                continue

            records_needing_search += 1
            keys: list[str] = []
            for term_item in stage2.get("suspect_terms", []) or []:
                term = normalize_text(
                    term_item.get("term")
                    if isinstance(term_item, dict) else term_item
                )
                if not term:
                    continue
                cache_key = make_search_cache_key(
                    record_id, term, record["original_text"]
                )
                keys.append(cache_key)
                cached = search_cache.get(cache_key, {})
                force_search = (
                    args.rerun_restored
                    and record_id in RESTORED_RECORD_IDS
                )
                if cached and not force_search and not (
                    cached.get("status") == "search_failed"
                    and args.rerun_stage2_failed
                ):
                    continue
                tasks.append(
                    {
                        "cache_key": cache_key,
                        "record_id": record_id,
                        "term": term,
                        "original_text": record["original_text"],
                        "stage1_code": stage1_map[record_id].get(
                            "first_level_code", ""
                        ),
                    }
                )
            task_keys_by_record[record_id] = keys

        print(
            f"[Stage 2批次启动 {label}] 本批{len(batch)}条；"
            f"需搜索记录={records_needing_search}条；"
            f"待请求术语={len(tasks)}个。",
            flush=True,
        )

        if tasks:
            search_started = time.perf_counter()
            completed = 0
            with ThreadPoolExecutor(
                max_workers=max(1, args.doubao_workers)
            ) as executor:
                future_map = {
                    executor.submit(search_one_with_retries, task, args): task
                    for task in tasks
                }
                unfinished = set(future_map)
                while unfinished:
                    done, unfinished = wait(
                        unfinished,
                        timeout=10,
                        return_when=FIRST_COMPLETED,
                    )
                    if not done:
                        print(
                            f"[等待Stage 2 {label}] "
                            f"已运行{time.perf_counter()-search_started:.0f}秒；"
                            f"完成{completed}/{len(tasks)}个术语；"
                            f"最多{min(max(1, args.doubao_workers), len(unfinished))}"
                            "个请求正在执行。",
                            flush=True,
                        )
                        continue
                    for future in done:
                        task = future_map[future]
                        try:
                            cache_key, search_result = future.result()
                        except Exception as exc:
                            cache_key = task["cache_key"]
                            search_result = {
                                "term": task["term"],
                                "meaning_cn": "",
                                "coding_relevance": "",
                                "confidence": "",
                                "needs_human_review": True,
                                "review_reason": "豆包联网搜索失败。",
                                "source_urls": [],
                                "status": "search_failed",
                                "error": f"{type(exc).__name__}: {exc}",
                            }
                        search_cache[cache_key] = search_result
                        completed += 1
                        save_search_cache()
                        elapsed = time.perf_counter() - search_started
                        percent = 100.0 * completed / len(tasks)
                        eta = elapsed / completed * (len(tasks) - completed)
                        print(
                            f"[Stage 2进度 {label} {percent:5.1f}%] "
                            f"{completed}/{len(tasks)}个术语；"
                            f"已用时={elapsed/60:.1f}分钟；"
                            f"预计剩余={eta/60:.1f}分钟。",
                            flush=True,
                        )

        # 将本批缓存搜索结果立即回填到记录。
        for record in batch:
            record_id = record["record_id"]
            stage2 = stage2_map.get(record_id, {})
            if not stage2.get("needs_search"):
                continue
            keys = task_keys_by_record.get(record_id, [])
            results = [search_cache[key] for key in keys if key in search_cache]
            stage2["search_results"] = results
            failures = [
                item for item in results
                if item.get("status") != "search_success"
            ]
            if len(results) == len(keys) and not failures and keys:
                stage2["status"] = "search_complete"
                stage2["error"] = ""
            else:
                stage2["status"] = "search_partial_failed"
                stage2["error"] = "；".join(
                    normalize_text(item.get("error")) for item in failures
                )[:2000]
            stage2_map[record_id] = stage2

        save_stage2_state()
        completed_records = sum(
            stage2_map.get(record["record_id"], {}).get("status")
            in {"no_search_needed", "search_complete"}
            for record in batch
        )
        print(
            f"[Stage 2批次完成 {label}] "
            f"本批已完成/无需搜索={completed_records}/{len(batch)}条。",
            flush=True,
        )
        # 搜索不是终点：本批搜索成功的记录立即回到Stage 3复核。
        review_searched_batch(batch, label)

    detection_pending = []
    for record in eligible_records:
        previous = stage2_map.get(record["record_id"], {})
        status = previous.get("status")
        if status in {"terms_detected", "no_search_needed", "search_complete", "search_partial_failed"}:
            continue
        if status == "term_detection_failed" and not args.rerun_stage2_failed:
            continue
        detection_pending.append(record)

    detection_batches = split_batches(detection_pending, max(1, args.stage2_batch_size))
    print(
        f"[Stage 1B] 已恢复{len(stage2_map)}条；"
        f"本次需识别{len(detection_pending)}条，共{len(detection_batches)}批。",
        flush=True,
    )

    # 如果上次在“识别完成、搜索开始前”中断，先给这些已识别记录补搜。
    recovered_detected = [
        record for record in eligible_records
        if stage2_map.get(record["record_id"], {}).get("status") == "terms_detected"
    ]
    recovered_batches = split_batches(
        recovered_detected, max(1, args.stage2_batch_size)
    )
    if recovered_batches:
        print(
            f"[断点接续] 有{len(recovered_detected)}条已完成识别但未搜索；"
            f"现在先分{len(recovered_batches)}批补做豆包搜索。",
            flush=True,
        )
        for recovered_index, recovered_batch in enumerate(
            recovered_batches, start=1
        ):
            search_detected_batch(
                recovered_batch,
                f"断点补搜{recovered_index}/{len(recovered_batches)}",
            )

    pipeline_started = time.perf_counter()
    for index, batch in enumerate(detection_batches, start=1):
        # 单批顺序处理：识别这一批，搜索这一批，再进入下一批。
        detect_started = time.perf_counter()
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(
                detect_terms_batch,
                records=batch,
                stage1_map=stage1_map,
                api_url=args.api_url,
                api_key=gemini_key,
                model=args.model,
                timeout=args.timeout,
                max_tokens=args.max_tokens,
                max_terms=args.max_search_terms,
                max_retries=args.max_retries,
            )
            while True:
                done, _ = wait(
                    {future}, timeout=10, return_when=FIRST_COMPLETED
                )
                if done:
                    detected = future.result()
                    break
                print(
                    f"[等待Stage 1B {index}/{len(detection_batches)}] "
                    f"本批已运行{time.perf_counter()-detect_started:.0f}秒；"
                    f"正在识别{len(batch)}条记录。",
                    flush=True,
                )
        stage2_map.update(detected)
        save_stage2_state()
        need_count = sum(
            item.get("needs_search", False)
            for item in stage2_map.values()
            if isinstance(item, dict)
        )
        elapsed = time.perf_counter() - pipeline_started
        percent = 100.0 * index / len(detection_batches)
        eta = elapsed / index * (len(detection_batches) - index)
        print(
            f"[Stage 1B进度 {percent:5.1f}%] "
            f"{index}/{len(detection_batches)}批；"
            f"当前累计需搜索={need_count}条；"
            f"已用时={elapsed/60:.1f}分钟；"
            f"预计剩余={eta/60:.1f}分钟。",
            flush=True,
        )
        search_detected_batch(batch, f"{index}/{len(detection_batches)}")

    # 空文本同步写入Stage 2，确保最终表行数完整。
    for record in records:
        if not record["original_text"]:
            stage2_map[record["record_id"]] = {
                "record_id": record["record_id"],
                "interpretation": "",
                "needs_search": False,
                "still_unclear": True,
                "uncertainty": "原文为空",
                "suspect_terms": [],
                "detection_memo": "原文为空。",
                "search_results": [],
                "status": "blank_text",
                "error": "",
            }

    stage1_template_path = export_stage1_template_file(
        records=records,
        stage1_map=stage1_map,
        stage2_map=stage2_map,
        output_dir=output_dir,
    )
    stage1_search_count = sum(
        bool(item.get("needs_search"))
        for item in stage2_map.values()
        if isinstance(item, dict)
    )
    print(
        f"[Stage 1完成] 需要豆包搜索={stage1_search_count}条；"
        f"已导出：{stage1_template_path}",
        flush=True,
    )

    # ---------- Stage 2：豆包Responses API + web_search ----------
    search_cache = load_checkpoint(search_cache_path)
    search_tasks = []
    all_task_keys: dict[str, list[str]] = {}
    for record in eligible_records:
        record_id = record["record_id"]
        stage2 = stage2_map.get(record_id, {})
        if not stage2.get("needs_search"):
            stage2["status"] = (
                "term_detection_failed"
                if stage2.get("status") == "term_detection_failed"
                else "no_search_needed"
            )
            stage2["search_results"] = []
            stage2_map[record_id] = stage2
            continue

        keys = []
        for term_item in stage2.get("suspect_terms", []) or []:
            term = normalize_text(term_item.get("term") if isinstance(term_item, dict) else term_item)
            if not term:
                continue
            cache_key = make_search_cache_key(record_id, term, record["original_text"])
            keys.append(cache_key)
            cached = search_cache.get(cache_key, {})
            if cached and not (
                cached.get("status") == "search_failed" and args.rerun_stage2_failed
            ):
                continue
            search_tasks.append(
                {
                    "cache_key": cache_key,
                    "record_id": record_id,
                    "term": term,
                    "original_text": record["original_text"],
                    "stage1_code": stage1_map[record_id].get("first_level_code", ""),
                }
            )
        all_task_keys[record_id] = keys

    print(
        f"[Stage 2] 检索缓存={len(search_cache)}项；"
        f"本次需要豆包联网搜索={len(search_tasks)}项。",
        flush=True,
    )

    if search_tasks:
        ark_key = os.getenv("ARK_API_KEY", "")
        if not ark_key:
            raise RuntimeError(
                "Stage 2需要豆包联网搜索，但.env中没有ARK_API_KEY。\n"
                "请配置：ARK_API_KEY=你的火山方舟API Key"
            )
        if not args.ark_model:
            raise RuntimeError(
                "Stage 2需要豆包联网搜索，但未配置ARK_MODEL。\n"
                "请在.env设置支持Responses API与web_search的模型或推理接入点。"
            )
        args.ark_api_key = ark_key
        started = time.perf_counter()
        completed_searches = 0
        with ThreadPoolExecutor(max_workers=max(1, args.doubao_workers)) as executor:
            future_map = {
                executor.submit(search_one_with_retries, task, args): task
                for task in search_tasks
            }
            unfinished = set(future_map)
            while unfinished:
                done, unfinished = wait(
                    unfinished,
                    timeout=10,
                    return_when=FIRST_COMPLETED,
                )
                if not done:
                    print(
                        f"[等待豆包搜索] 已运行{time.perf_counter()-started:.0f}秒；"
                        f"完成{completed_searches}/{len(search_tasks)}项。",
                        flush=True,
                    )
                    continue
                for future in done:
                    cache_key, search_result = future.result()
                    search_cache[cache_key] = search_result
                    completed_searches += 1
                    atomic_write_json(
                        search_cache_path,
                        {
                            "stage": "doubao_web_search_cache",
                            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "results": search_cache,
                        },
                    )
                    succeeded = sum(
                        item.get("status") == "search_success"
                        for item in search_cache.values()
                        if isinstance(item, dict)
                    )
                    print(
                        f"[豆包搜索进度] {completed_searches}/{len(search_tasks)}项；"
                        f"缓存中成功={succeeded}。",
                        flush=True,
                    )

    # 将搜索缓存回填到每条记录的Stage 2结果。
    for record in eligible_records:
        record_id = record["record_id"]
        stage2 = stage2_map.get(record_id, {})
        if not stage2.get("needs_search"):
            continue
        results = [
            search_cache[key]
            for key in all_task_keys.get(record_id, [])
            if key in search_cache
        ]
        stage2["search_results"] = results
        failures = [item for item in results if item.get("status") != "search_success"]
        expected = len(all_task_keys.get(record_id, []))
        if len(results) == expected and not failures:
            stage2["status"] = "search_complete"
            stage2["error"] = ""
        else:
            stage2["status"] = "search_partial_failed"
            stage2["error"] = "；".join(
                normalize_text(item.get("error")) for item in failures
            )[:2000]
        stage2_map[record_id] = stage2

    atomic_write_json(
        stage2_checkpoint,
        {
            "stage": "stage2_term_detection_and_doubao_search",
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "results": stage2_map,
        },
    )
    searched_records = sum(
        item.get("status") == "search_complete"
        for item in stage2_map.values()
        if isinstance(item, dict)
    )
    print(f"[Stage 2完成] 成功搜索并回填={searched_records}条。", flush=True)

    # ---------- Stage 3：搜索后复核与回填 ----------
    print("\n========== Stage 3：基于搜索结果复核与回填 ==========", flush=True)
    stage3_map = load_checkpoint(stage3_checkpoint)
    # --rerun-restored 已在Stage 1B开始前清除了旧Stage 3结果；
    # 不要在这里再次清除，否则会删除本轮“搜索后立即复核”的新结果，
    # 导致同一记录被Gemini重复复核。
    review_pending = []
    for record in eligible_records:
        record_id = record["record_id"]
        stage1 = stage1_map[record_id]
        stage2 = stage2_map.get(record_id, {})
        previous = stage3_map.get(record_id, {})
        previous_status = previous.get("status")

        # 只有“已经基于搜索结果复核成功”才可直接恢复。
        # 旧checkpoint中的copied_no_search不能覆盖后来新增的搜索结果。
        if (
            stage2.get("needs_search")
            and stage2.get("status") == "search_complete"
        ):
            if previous_status == "stage3_success":
                continue
            review_pending.append(record)
            continue

        if previous_status == "copied_no_search" and not stage2.get("needs_search"):
            continue
        if (
            previous_status in {"stage3_failed", "copied_search_failed"}
            and stage2.get("status") != "search_complete"
            and not args.rerun_stage3_failed
        ):
            continue
        if not stage2.get("needs_search"):
            detection_failed = stage2.get("status") == "term_detection_failed"
            stage3_map[record_id] = make_stage3_copy(
                stage1,
                status="copied_search_failed" if detection_failed else "copied_no_search",
                review=detection_failed,
                reason=(
                    "Stage 2术语识别失败，暂时保留Stage 1编码并需人工复核。"
                    if detection_failed else ""
                ),
                interpretation=stage2.get("interpretation", ""),
                still_unclear=stage2.get("still_unclear", False),
                uncertainty=stage2.get("uncertainty", ""),
            )
        elif stage2.get("status") != "search_complete":
            stage3_map[record_id] = make_stage3_copy(
                stage1,
                status="copied_search_failed",
                review=True,
                reason="豆包搜索未完整成功，暂时保留Stage 1编码并需人工复核。",
                interpretation=stage2.get("interpretation", ""),
                still_unclear=True,
                uncertainty="豆包搜索未完整成功",
            )
        else:
            # 理论上search_complete已在上方加入复核队列；
            # 保留此分支作为异常状态的安全补充。
            review_pending.append(record)

    stage3_batches = split_batches(review_pending, max(1, args.stage3_batch_size))
    print(
        f"[Stage 3] checkpoint已有={len(stage3_map)}条；"
        f"其中需基于新搜索结果调用Gemini复核="
        f"{len(review_pending)}条，共{len(stage3_batches)}批。",
        flush=True,
    )
    stage3_started = time.perf_counter()
    for index, batch in enumerate(stage3_batches, start=1):
        batch_started = time.perf_counter()
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(
                review_stage3_batch,
                records=batch,
                stage1_map=stage1_map,
                stage2_map=stage2_map,
                api_url=args.api_url,
                api_key=gemini_key,
                model=args.model,
                timeout=args.timeout,
                max_tokens=args.max_tokens,
                max_retries=args.max_retries,
            )
            while True:
                done, _ = wait(
                    {future}, timeout=10, return_when=FIRST_COMPLETED
                )
                if done:
                    reviewed = future.result()
                    break
                print(
                    f"[等待Stage 3 {index}/{len(stage3_batches)}] "
                    f"本批已运行{time.perf_counter()-batch_started:.0f}秒；"
                    f"正在复核{len(batch)}条记录。",
                    flush=True,
                )
        stage3_map.update(reviewed)
        atomic_write_json(
            stage3_checkpoint,
            {
                "stage": "stage3_search_informed_review",
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "results": stage3_map,
            },
        )
        elapsed = time.perf_counter() - stage3_started
        percent = 100.0 * index / len(stage3_batches)
        eta = elapsed / index * (len(stage3_batches) - index)
        print(
            f"[Stage 3进度 {percent:5.1f}%] "
            f"{index}/{len(stage3_batches)}批；"
            f"已用时={elapsed/60:.1f}分钟；"
            f"预计剩余={eta/60:.1f}分钟。",
            flush=True,
        )

    for record in records:
        if not record["original_text"]:
            stage3_map[record["record_id"]] = {
                "record_id": record["record_id"],
                "final_first_level_code": "",
                "changed_from_stage1": "",
                "evidence": "",
                "memo": "原文为空。",
                "confidence": "",
                "needs_review": True,
                "review_reason": "请回查原始数据。",
                "status": "blank_text",
                "error": "",
            }

    atomic_write_json(
        stage3_checkpoint,
        {
            "stage": "stage3_search_informed_review",
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "results": stage3_map,
        },
    )
    stage1_path, stage2_path, stage3_path, json_path = export_stage_template_files(
        records=records,
        stage1_map=stage1_map,
        stage2_map=stage2_map,
        stage3_map=stage3_map,
        output_dir=output_dir,
    )
    print("\nStage 1→豆包搜索→Stage 3完整流程结束", flush=True)
    print(f"Stage 1结果：{stage1_path}", flush=True)
    print(f"Stage 2结果：{stage2_path}", flush=True)
    print(f"Stage 3结果：{stage3_path}", flush=True)
    print(f"完整JSON：{json_path}", flush=True)
    print(f"Stage 2断点：{stage2_checkpoint}", flush=True)
    print(f"Stage 3断点：{stage3_checkpoint}", flush=True)


if __name__ == "__main__":
    main_full_pipeline()
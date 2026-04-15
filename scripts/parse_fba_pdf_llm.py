#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parse a single Amazon FBA / BOXSTAR shipment PDF using OpenRouter (GPT-5.4).

Each page of the PDF is a single-box label. We extract the raw text of every
page with pdfplumber, send the text (not an image) to the LLM in parallel
(max 20 concurrent), and ask the model to extract the structured fields.
Results are aggregated into the JSON schema consumed by
scripts/generate_all.py: { shipments: [...], matrix: {...} }.

Text-in is ~3-4x cheaper than image-in for these labels and avoids OCR
mistakes on CJK addresses / SKUs.

Dependencies:
    pip install openai --break-system-packages -q
    # pdftotext is from poppler (brew install poppler)
"""

import argparse
import json
import os
import random
import re
import shutil
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional

# --- Config ---------------------------------------------------------------
OPENROUTER_API_KEY = "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"  # placeholder; prefer env var
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
MODEL = "openai/gpt-5.4"
MAX_CONCURRENCY = 20
MAX_RETRIES = 6           # retries on rate-limit / transient failures
RETRY_BASE_DELAY = 2.0    # seconds; exponential backoff with jitter
REQUEST_TIMEOUT = 120     # seconds per request

SYSTEM_PROMPT = """你是识别跨境电商分仓货件箱单标签的专家。
用户给你**一页箱单标签的纯文本**（由 poppler 的 `pdftotext -layout` 抽出，保留了原始的版面对齐——**同一行里的多栏内容用大段空格水平分开**，栏目边界视觉上清晰可辨）。这一页对应**一个纸箱**。同时用户会给你该 PDF 的**文件名**（文件名里有时包含仓库代码，见下文）。
你必须只输出一个 JSON 对象，不要解释、不要 markdown 代码块。

============================================================
Part 1. 通用逻辑——每种标签都要找的 6 个字段
============================================================

不论标签长什么样，我们需要的是这六项信息：

- `box_number` (int)：当前箱的序号（第几箱）。
- `total_boxes` (int)：本票一共多少箱。通常和 box_number 出现在同一行，形如 "N / M"、"N of M"、"N，共 M 个"。
- `warehouse_code` (str | null)：目的仓库的代码。**只有在文本里或文件名里明确出现时才填**。找不到就填 null，**绝不从地址/城市/邮编推断**。
- `address` (str | null)：收货地址（通常是美国地址）。有就拼成一行（用英文逗号+空格分隔），没有就 null。
- `sku` (str)：这一箱装的 SKU。SKU 通常是**全大写字母+数字，可能带 `-` 或小写后缀**，例如 `PTR220001-P`、`BPBBB25000MS`、`BPBBB25000LS-a`。
- `qty_per_box` (int)：**这一个箱子里**该 SKU 的件数（不是整票总数！）。

============================================================
Part 2. 已见过的版面（基于抽出的纯文本）
============================================================

ℹ️ 提示：由于使用了 `pdftotext -layout`，页面中**左右并排的两栏**（"目的地"栏 vs "发货地"栏）在文本里会处于同一行但被**大段空格**隔开。你可以把每行按视觉位置拆成左半段和右半段——只关心**左半段的"目的地"内容**，忽略右半段里的 `shenzhenshi tangmumao...` / `Guangdong - 深圳 ...` / `南山街道...` / `中国` 等中国发货方内容。

▶ 版面 A：亚马逊 FBA 标签
  抽出的文本大致长这样（示例，空格原样保留以保持对齐）：
    ```
    FBA                                                       纸箱编号 1，共 14 个纸箱 - 34 磅
    目的地：                                                     发货地：
    FBA: shenzhenshi tangmumao dian shang you xian gong si   shenzhenshi tangmumao dian shang you xian
    MDW2                                                     Guangdong - 深圳 - 518052
    250 EMERALD DR                                           南山街道荔湾社区前海路0101号丽湾商务公寓A-2519
    Joliet, IL 60433-3280                                    中国
    美国
    塑料稻草-146                                                             Created: 2026/03/24 01:10 CDT (-05)

                        FBA199DN1ZLPU000001

                                                                  Single SKU
                                                                  PTR220001-P
                                                                      数量 25
                                                                         A-25

                                              请不要遮住此标签
    ```
  - 第 1 行 `纸箱编号 N，共 M 个纸箱` → `box_number=N`, `total_boxes=M`。
  - `FBA: shenzhenshi ...` 那一行（左半段）是卖家名，忽略。
  - 紧接着（目的地块第 2 行）有两种形式：
    · **A1**：该行**左半段**就是一个独立的 3–4 位大写字母/数字代码（如 `MDW2`），右半段是 `Guangdong - 深圳 ...`（忽略）。这就是 `warehouse_code`。下面几行的左半段是美国地址。
    · **A2**：该行**左半段**是 `Amazon.com Services, Inc.` / `Amazon.com Services LLC` 之类的公司名，标签上**没有**仓库代码行。这时从**文件名**里取仓库代码——FBA 文件名通常形如 `FBA199DNP1K1-AVP1.pdf`，末尾 `-XXXX.pdf` 之前的部分（`AVP1`）就是仓库代码。能取到就用它，否则 `warehouse_code=null`。
  - 美国地址：紧跟在"仓库代码行"（A1）或"Amazon 公司名行"（A2）之后的 2 行的**左半段**，拼成一行，例如：
      A1 → "250 EMERALD DR, Joliet, IL 60433-3280"
      A2 → "550 Oak Ridge Road, Hazle Township, PA 18202-9361"
    不要把 "FBA: ..." 那行、仓库代码、"Amazon.com Services" 公司名、"美国" 单独一行放进地址。
  - SKU / 数量：在 `Single SKU` 之后紧跟着的那行就是 `sku`（如 `PTR220001-P`），再下一行 `数量 X` → `qty_per_box=X`，再下面 `A-25` 是库位，忽略。
  - 忽略的其它内容：`塑料稻草-xxx`、`Created: ...`、`FBA...U000001`（条码下方长串）、`请不要遮住此标签`、右上角 `XX磅`、右半段所有发货地内容。

▶ 版面 B：BOXSTAR（置闰）标签
  抽出的文本大致长这样（`-layout` 保留了表格列对齐）：
    ```
    Inbound                              Box 1 of 120

    WH: BOX STAR (AFDLA01)               SKUs: 1
    Client: 置闰(7538006)                  PCS: 1

    SKU                      Item Name         Qty
    BPBBB25000LS-a           黑色宠物床边床大号 1

    IB006260408RT                        2026-04-08 10:09

                    BPBBB25000LS-a           MADE IN CHINA
    ```
  - `Box N of M` → `box_number=N`, `total_boxes=M`。
  - `WH: BOX STAR (XXXXXX)`：**括号里的代码就是 `warehouse_code`**（如 `AFDLA01`），原样大写输出。
  - `Client: 置闰(...)`、`SKUs: 1`、`PCS: 1`：忽略。
  - `SKU Item Name Qty` 是表头，下一行是数据行：第一个 token 是 `sku`（如 `BPBBB25000MS`、`BPBBB25000LS-a`），**最后一个** token 是 `qty_per_box`（整数）。中间的中文品名（如 `黑色宠物床边床大号`）忽略。
  - BOXSTAR 标签**没有美国地址**，`address=null`。
  - 忽略：`IB...RT` 编号、时间戳、第二次重复出现的 SKU、`MADE IN CHINA`。

============================================================
Part 3. 没见过的版面
============================================================

如果这张标签不是上面任何一种，按 Part 1 的通用逻辑尽力识别。找不到某字段就填 null，**不要编造**，不要从无关内容推断仓库代码。

输出 JSON 示例：
{"box_number": 1, "total_boxes": 14, "warehouse_code": "MDW2", "address": "250 EMERALD DR, Joliet, IL 60433-3280", "sku": "PTR220001-P", "qty_per_box": 25}
{"box_number": 1, "total_boxes": 14, "warehouse_code": "AVP1", "address": "550 Oak Ridge Road, Hazle Township, PA 18202-9361", "sku": "PTR220001-P", "qty_per_box": 25}
{"box_number": 1, "total_boxes": 120, "warehouse_code": "AFDLA01", "address": null, "sku": "BPBBB25000LS-a", "qty_per_box": 1}
"""

USER_PROMPT_TEMPLATE = """PDF 文件名：{filename}

以下是这一页标签的纯文本（pdfplumber 抽取，原始行顺序保留）：
----- BEGIN PAGE TEXT -----
{page_text}
----- END PAGE TEXT -----

请按系统指令的 JSON 格式输出。"""


# --- PDF text extraction --------------------------------------------------
def extract_pdf_pages_text(pdf_path: str) -> List[str]:
    """Extract per-page text via `pdftotext -layout` (poppler).

    Uses reading-order + column-aligned whitespace, which matches what a user
    gets when selecting-and-copying text in Preview/Chrome. Columns stay
    visually separated (the destination block and sender block are no longer
    glued onto the same line).

    Pages are split on the form-feed character (\\f), which pdftotext emits
    between pages.
    """
    if shutil.which("pdftotext") is None:
        print("Error: pdftotext not found. Install poppler: brew install poppler",
              file=sys.stderr)
        sys.exit(1)

    try:
        proc = subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8", pdf_path, "-"],
            capture_output=True,
            check=True,
        )
    except subprocess.CalledProcessError as e:
        print(f"Error: pdftotext failed on {pdf_path}: {e.stderr.decode('utf-8', errors='replace')}",
              file=sys.stderr)
        sys.exit(1)

    raw = proc.stdout.decode("utf-8", errors="replace")
    # pdftotext emits a trailing \f; strip it so we don't get an empty page.
    pages = raw.split("\f")
    if pages and pages[-1].strip() == "":
        pages = pages[:-1]
    return pages


# --- LLM call -------------------------------------------------------------
def _get_client(api_key: Optional[str]):
    try:
        from openai import OpenAI
    except ImportError:
        print("Error: openai SDK not installed. Run: pip install openai --break-system-packages",
              file=sys.stderr)
        sys.exit(1)

    key = api_key or os.environ.get("OPENROUTER_API_KEY") or OPENROUTER_API_KEY
    if not key or key == "REPLACE_WITH_OPENROUTER_API_KEY":
        print("Error: no OpenRouter API key. Set OPENROUTER_API_KEY env var or pass --api-key.",
              file=sys.stderr)
        sys.exit(1)
    return OpenAI(base_url=OPENROUTER_BASE_URL, api_key=key)


def _is_rate_limit_error(err: Exception) -> bool:
    msg = str(err).lower()
    if "rate" in msg and "limit" in msg:
        return True
    if "429" in msg or "too many requests" in msg:
        return True
    # Also retry on 5xx / timeouts / generic transient
    for marker in ("timeout", "timed out", "temporarily", "502", "503", "504", "overload"):
        if marker in msg:
            return True
    return False


def _extract_json(text: str) -> Dict:
    """Extract a JSON object from the model's response, tolerant of code fences."""
    text = text.strip()
    # Strip ```json ... ``` fences if present
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    # Try direct parse first
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # Fallback: grab the first {...} block
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        raise ValueError(f"No JSON object found in model output: {text[:200]!r}")
    return json.loads(m.group(0))


def call_llm_for_page(client, page_index: int, page_text: str, filename: str = "") -> Dict:
    """Call OpenRouter GPT-5.4 with the page's extracted text. Retries on rate limits."""
    user_msg = USER_PROMPT_TEMPLATE.format(filename=filename, page_text=page_text)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_msg},
    ]

    last_err: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=messages,
                response_format={"type": "json_object"},
                timeout=REQUEST_TIMEOUT,
            )
            content = resp.choices[0].message.content or ""
            parsed = _extract_json(content)
            parsed["_page_index"] = page_index  # 0-based, for ordering/debug
            return parsed
        except Exception as e:  # noqa: BLE001
            last_err = e
            if attempt >= MAX_RETRIES or not _is_rate_limit_error(e):
                # Non-retryable or out of retries
                if attempt >= MAX_RETRIES:
                    break
                raise
            delay = RETRY_BASE_DELAY * (2 ** (attempt - 1)) + random.uniform(0, 1.0)
            print(f"  [page {page_index + 1}] transient error ({e}); retry {attempt}/{MAX_RETRIES} in {delay:.1f}s",
                  file=sys.stderr)
            time.sleep(delay)

    raise RuntimeError(f"Page {page_index + 1} failed after {MAX_RETRIES} retries: {last_err}")


# --- Aggregation ----------------------------------------------------------
def _majority(values: List) -> Optional[object]:
    """Return the most common non-null value."""
    from collections import Counter
    cleaned = [v for v in values if v not in (None, "", 0)]
    if not cleaned:
        return None
    return Counter(cleaned).most_common(1)[0][0]


def aggregate_sku_breakdown(pages: List[Dict]) -> List[Dict]:
    """Group consecutive boxes with same SKU+qty_per_box into one breakdown row.

    Pages are expected to already be sorted by box_number.
    """
    breakdown: List[Dict] = []
    current_sku = None
    current_qty = None
    box_count = 0

    for p in pages:
        sku = p.get("sku")
        qty = p.get("qty_per_box")
        if not sku or not qty:
            continue
        if sku == current_sku and qty == current_qty:
            box_count += 1
        else:
            if current_sku is not None:
                breakdown.append({
                    "sku": current_sku,
                    "boxes": box_count,
                    "qty_per_box": current_qty,
                    "total_qty": box_count * current_qty,
                })
            current_sku = sku
            current_qty = qty
            box_count = 1

    if current_sku is not None:
        breakdown.append({
            "sku": current_sku,
            "boxes": box_count,
            "qty_per_box": current_qty,
            "total_qty": box_count * current_qty,
        })

    return breakdown


def build_shipment(pdf_path: str, page_results: List[Dict]) -> Dict:
    """Aggregate per-page LLM results into the shipment JSON structure."""
    filename = os.path.basename(pdf_path)

    # Sort by box_number if present, else by page index
    def sort_key(p):
        bn = p.get("box_number")
        return (bn if isinstance(bn, int) else p.get("_page_index", 0))
    page_results = sorted(page_results, key=sort_key)

    warehouse_code = _majority([p.get("warehouse_code") for p in page_results])
    address = _majority([p.get("address") for p in page_results])
    total_boxes = _majority([p.get("total_boxes") for p in page_results]) or len(page_results)

    sku_breakdown = aggregate_sku_breakdown(page_results)

    # Expose per-page raw LLM results for auditing. Strip internal keys and
    # renumber page_number to be 1-based.
    pages_debug = []
    for p in page_results:
        clean = {k: v for k, v in p.items() if not k.startswith("_")}
        clean["page_number"] = p.get("_page_index", 0) + 1
        pages_debug.append(clean)

    return {
        "file": filename,
        "warehouse_code": warehouse_code,
        "address": address or "",
        "total_boxes": int(total_boxes) if total_boxes else len(page_results),
        "sku_breakdown": sku_breakdown,
        "pages": pages_debug,
    }


# --- Main -----------------------------------------------------------------
def parse_pdf_with_llm(pdf_path: str, api_key: Optional[str] = None,
                       concurrency: int = MAX_CONCURRENCY) -> Dict:
    concurrency = max(1, min(concurrency, MAX_CONCURRENCY))
    print(f"Extracting PDF text: {pdf_path}", file=sys.stderr)
    pages_text = extract_pdf_pages_text(pdf_path)
    print(f"  {len(pages_text)} page(s); dispatching to {MODEL} (max {concurrency} parallel)",
          file=sys.stderr)

    client = _get_client(api_key)
    results: List[Dict] = []

    filename = os.path.basename(pdf_path)
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = {
            pool.submit(call_llm_for_page, client, i, txt, filename): i
            for i, txt in enumerate(pages_text)
        }
        for fut in as_completed(futures):
            page_idx = futures[fut]
            try:
                results.append(fut.result())
                print(f"  [page {page_idx + 1}] ok", file=sys.stderr)
            except Exception as e:  # noqa: BLE001
                print(f"  [page {page_idx + 1}] FAILED: {e}", file=sys.stderr)
                raise

    return build_shipment(pdf_path, results)


def _build_matrix(shipments: List[Dict]) -> Dict[str, Dict[str, int]]:
    matrix: Dict[str, Dict[str, int]] = {}
    for s in shipments:
        wh = s.get("warehouse_code") or "UNKNOWN"
        for item in s.get("sku_breakdown", []):
            matrix.setdefault(item["sku"], {})[wh] = (
                matrix.setdefault(item["sku"], {}).get(wh, 0) + item["total_qty"]
            )
    return matrix


def main():
    ap = argparse.ArgumentParser(
        description="Parse FBA / BOXSTAR shipment PDF(s) via OpenRouter GPT-5.4 "
                    "(text-based; pdfplumber extracts per-page text, LLM extracts fields). "
                    "Accepts a single PDF or a folder containing multiple PDFs."
    )
    ap.add_argument("input_path", help="Single FBA shipment PDF, or a folder of PDFs")
    ap.add_argument("--output", default="/tmp/fba_shipments.json",
                    help="Output JSON path (default: /tmp/fba_shipments.json)")
    ap.add_argument("--api-key", default=None,
                    help="OpenRouter API key (overrides OPENROUTER_API_KEY env var)")
    ap.add_argument("--concurrency", type=int, default=MAX_CONCURRENCY,
                    help=f"Max parallel requests per PDF (default/ceiling: {MAX_CONCURRENCY}). "
                         "PDFs are processed sequentially to avoid rate-limit storms.")
    args = ap.parse_args()

    # Collect PDF(s)
    input_path = Path(args.input_path)
    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            print(f"Error: {args.input_path} is not a PDF file", file=sys.stderr)
            sys.exit(1)
        pdf_files = [input_path]
    elif input_path.is_dir():
        pdf_files = sorted(input_path.glob("*.pdf"))
        if not pdf_files:
            print(f"Error: no PDF files found in {args.input_path}", file=sys.stderr)
            sys.exit(1)
    else:
        print(f"Error: {args.input_path} does not exist", file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(pdf_files)} PDF(s) to parse", file=sys.stderr)

    # Process PDFs sequentially; pages within each PDF are parallelized
    shipments: List[Dict] = []
    for i, pdf in enumerate(pdf_files, 1):
        print(f"\n=== [{i}/{len(pdf_files)}] {pdf.name} ===", file=sys.stderr)
        shipment = parse_pdf_with_llm(str(pdf), api_key=args.api_key,
                                      concurrency=args.concurrency)
        shipments.append(shipment)

    output = {"shipments": shipments, "matrix": _build_matrix(shipments)}

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"\nOutput written to {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parse a single Amazon FBA shipment PDF using OpenRouter (GPT-5.4) vision.

Each page of the PDF is a single-box label. We render every page to a PNG,
send it to the LLM in parallel (max 20 concurrent), and ask the model to
extract the structured fields. Results are aggregated into the JSON schema
consumed by scripts/generate_all.py: { shipments: [...], matrix: {...} }.

Dependencies:
    pip install openai pdf2image Pillow --break-system-packages -q
    # pdf2image requires poppler (brew install poppler)
"""

import argparse
import base64
import io
import json
import os
import random
import re
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
RENDER_DPI = 200          # rasterization resolution

SYSTEM_PROMPT = """你是识别跨境电商分仓货件箱单标签的专家。
用户给你一张 PDF 单页截图，这一页对应**一个纸箱**的发货标签，同时会告诉你该 PDF 的**文件名**（文件名里有时包含仓库代码，见下文）。
你必须只输出一个 JSON 对象，不要解释、不要 markdown 代码块。

============================================================
Part 1. 通用逻辑——每种标签都要找的 6 个字段
============================================================

不论标签长什么样，我们需要的是这六项信息：

- `box_number` (int)：当前箱的序号（第几箱）。
- `total_boxes` (int)：本票一共多少箱。通常和 box_number 出现在同一行，形如 "N / M"、"N of M"、"N，共 M 个"。
- `warehouse_code` (str | null)：目的仓库的代码。**只有在标签本身或文件名里明确出现时才填**。找不到就填 null，**绝不从地址/城市/邮编推断**。
- `address` (str | null)：收货地址（通常是美国地址）。有就拼成一行（用英文逗号+空格分隔），没有就 null。
- `sku` (str)：这一箱装的 SKU。SKU 通常是**全大写字母+数字，可能带 `-` 或小写后缀**，例如 `PTR220001-P`、`BPBBB25000MS`、`BPBBB25000LS-a`。SKU 一般会在条码下方再次出现，可用于交叉验证。
- `qty_per_box` (int)：**这一个箱子里**该 SKU 的件数（不是整票总数！）。

============================================================
Part 2. 已见过的版面
============================================================

▶ 版面 A：亚马逊 FBA 标签
  - 页眉左上有大号 `FBA` logo；右上一行 `纸箱编号 N，共 M 个纸箱 - XX磅`
    → `box_number = N`，`total_boxes = M`。
  - 左侧"目的地:"下第一行是 `FBA: shenzhenshi tangmumao dian shang you xian gong si`。
  - 紧跟着有两种形式：
    · **A1**：下一行是一个独立的 3–4 位大写字母/数字代码（如 `MDW2`），这就是 `warehouse_code`，再下面才是美国地址。
    · **A2**：下一行直接是 `Amazon.com Services, Inc.` / `Amazon.com Services LLC` 之类的公司名，标签上**没有**仓库代码。
      这种情况下请**查看用户给的文件名**——FBA 文件名通常形如 `FBA199DNP1K1-AVP1.pdf`，末尾 `-XXXX.pdf` 之前的部分就是仓库代码。
      如果文件名里能提取出这个代码（全大写字母/数字，3–4 位）就用它；否则 `warehouse_code = null`。
  - 美国地址通常是 3–4 行（街道/城市、州 邮编/"美国"），拼成一行；**不要**包含 "FBA:" 那行、A1 型的仓库代码行、公司名行、"美国" 那行。
    示例："250 EMERALD DR, Joliet, IL 60433-3280" 或 "550 Oak Ridge Road, Hazle Township, PA 18202-9361"
  - 中部条码下方有 "Single SKU" 小块：下一行是 `sku`（如 `PTR220001-P`），再下一行 `数量 X` → `qty_per_box = X`，再下面的 `A-25` 是库位，忽略。
  - 忽略：`FBA` logo、发货地块（`Guangdong - 深圳 ...` / 南山街道 ... / 中国）、条码本身、条码下方的 `FBA...U000001` 长串、`Created: ...` 时间戳、`请不要遮住此标签`、左上 `纸箱指纹-xxx`、右上角 `XX磅`。

▶ 版面 B：BOXSTAR（置闰）标签
  - 页眉左上 `Inbound`，右上 `Box N of M` → `box_number=N`, `total_boxes=M`。
  - 下面一行 `WH: BOX STAR (XXXXXX)`，**括号里的代码就是 `warehouse_code`**（如 `AFDLA01`），原样大写输出。
  - 再下一行 `Client: 置闰(7538006)` —— 忽略。
  - 右上 `SKUs: 1` / `PCS: 1` —— 忽略。
  - 中部一个 3 列小表格，表头 `SKU | Item Name | Qty`：
    · `SKU` 列的值 → `sku`（如 `BPBBB25000MS`、`BPBBB25000LS-a`）。
    · `Item Name` 列是中文品名（如 `黑色宠物床边床中号`）—— 我们**不需要**，忽略。
    · `Qty` 列的值 → `qty_per_box`。
  - BOXSTAR 标签**没有美国地址**，`address = null`。
  - 忽略：下半部分的 `IB...RU` 编号和时间戳、第二个条码及其下方重复的 SKU、`MADE IN CHINA`、二维码。

============================================================
Part 3. 没见过的版面
============================================================

如果这张标签不是上面任何一种，按 Part 1 的通用逻辑尽力识别。找不到某字段就填 null，**不要编造**，不要从无关内容推断仓库代码。

输出 JSON 示例：
{"box_number": 1, "total_boxes": 14, "warehouse_code": "MDW2", "address": "250 EMERALD DR, Joliet, IL 60433-3280", "sku": "PTR220001-P", "qty_per_box": 25}
{"box_number": 1, "total_boxes": 141, "warehouse_code": "AFDLA01", "address": null, "sku": "BPBBB25000MS", "qty_per_box": 1}
"""

USER_PROMPT_TEMPLATE = "PDF 文件名：{filename}\n请识别这一页箱单标签，按系统指令的 JSON 格式输出。"


# --- PDF rendering --------------------------------------------------------
def render_pdf_pages_to_png(pdf_path: str, dpi: int = RENDER_DPI) -> List[bytes]:
    """Render every page of a PDF to PNG bytes using pdf2image (poppler)."""
    try:
        from pdf2image import convert_from_path
    except ImportError:
        print("Error: pdf2image not installed. Run: pip install pdf2image Pillow --break-system-packages",
              file=sys.stderr)
        sys.exit(1)

    images = convert_from_path(pdf_path, dpi=dpi)
    out = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        out.append(buf.getvalue())
    return out


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


def call_llm_for_page(client, page_index: int, png_bytes: bytes, filename: str = "") -> Dict:
    """Call OpenRouter GPT-5.4 with one page image. Retries on rate limits."""
    b64 = base64.b64encode(png_bytes).decode("ascii")
    data_url = f"data:image/png;base64,{b64}"

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": USER_PROMPT_TEMPLATE.format(filename=filename)},
                {"type": "image_url", "image_url": {"url": data_url}},
            ],
        },
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
    print(f"Rendering PDF pages: {pdf_path}", file=sys.stderr)
    pages_png = render_pdf_pages_to_png(pdf_path)
    print(f"  {len(pages_png)} page(s); dispatching to {MODEL} (max {concurrency} parallel)",
          file=sys.stderr)

    client = _get_client(api_key)
    results: List[Dict] = []

    filename = os.path.basename(pdf_path)
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = {
            pool.submit(call_llm_for_page, client, i, png, filename): i
            for i, png in enumerate(pages_png)
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
        description="Parse FBA shipment PDF(s) via OpenRouter GPT-5.4 vision. "
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

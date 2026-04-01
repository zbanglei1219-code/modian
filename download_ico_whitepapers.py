#!/usr/bin/env python3
"""Download ICO whitepaper PDFs listed in an Excel file.

Features:
- Reads token + whitepaper-link columns from .xlsx.
- Saves files as <token>.pdf (sanitized).
- Writes a detailed report for each row.
- Retries failed direct downloads through the Wayback Machine.
"""

from __future__ import annotations

import argparse
import csv
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
import requests


DEFAULT_TIMEOUT = 30
WAYBACK_CDX = "https://web.archive.org/cdx/search/cdx"


@dataclass
class DownloadResult:
    ok: bool
    status: str
    status_code: Optional[int]
    source_url: str
    saved_path: Optional[Path] = None
    error: Optional[str] = None
    via_wayback: bool = False
    wayback_snapshot: Optional[str] = None


def sanitize_filename(name: str, fallback: str) -> str:
    cleaned = (name or "").strip()
    if not cleaned:
        cleaned = fallback
    cleaned = cleaned.replace("/", "_")
    cleaned = re.sub(r"[\\:*?\"<>|\r\n\t]+", "_", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if not cleaned:
        cleaned = fallback
    return cleaned[:200]


def looks_like_pdf(content: bytes, content_type: str, url: str) -> bool:
    head = content[:8]
    if head.startswith(b"%PDF-"):
        return True
    ct = (content_type or "").lower()
    if "application/pdf" in ct:
        return True
    if url.lower().split("?")[0].endswith(".pdf"):
        return True
    return False


def download_pdf(session: requests.Session, url: str, out_path: Path, timeout: int) -> DownloadResult:
    try:
        resp = session.get(url, timeout=timeout, allow_redirects=True)
    except requests.RequestException as exc:
        return DownloadResult(
            ok=False,
            status="direct_request_exception",
            status_code=None,
            source_url=url,
            error=str(exc),
        )

    status_code = resp.status_code
    if status_code >= 400:
        return DownloadResult(
            ok=False,
            status="direct_http_error",
            status_code=status_code,
            source_url=url,
            error=f"HTTP {status_code}",
        )

    content = resp.content
    if not content:
        return DownloadResult(
            ok=False,
            status="direct_empty_content",
            status_code=status_code,
            source_url=str(resp.url),
            error="empty response body",
        )

    if not looks_like_pdf(content, resp.headers.get("content-type", ""), str(resp.url)):
        return DownloadResult(
            ok=False,
            status="direct_not_pdf",
            status_code=status_code,
            source_url=str(resp.url),
            error=f"content-type={resp.headers.get('content-type', '')}",
        )

    out_path.write_bytes(content)
    return DownloadResult(
        ok=True,
        status="downloaded_direct",
        status_code=status_code,
        source_url=str(resp.url),
        saved_path=out_path,
    )


def find_wayback_snapshot(session: requests.Session, original_url: str, timeout: int) -> Optional[str]:
    params = {
        "url": original_url,
        "output": "json",
        "fl": "timestamp,original,statuscode,mimetype",
        "filter": "statuscode:200",
        "collapse": "digest",
    }

    try:
        resp = session.get(WAYBACK_CDX, params=params, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
    except Exception:
        return None

    if not isinstance(data, list) or len(data) < 2:
        return None

    # data[0] is header row, choose latest valid snapshot
    rows = data[1:]
    rows.sort(key=lambda r: r[0], reverse=True)
    for row in rows:
        if len(row) < 2:
            continue
        ts = row[0]
        orig = row[1]
        snapshot_url = f"https://web.archive.org/web/{ts}id_/{orig}"
        return snapshot_url
    return None


def download_via_wayback(session: requests.Session, original_url: str, out_path: Path, timeout: int) -> DownloadResult:
    snapshot = find_wayback_snapshot(session, original_url, timeout)
    if not snapshot:
        return DownloadResult(
            ok=False,
            status="wayback_no_snapshot",
            status_code=None,
            source_url=original_url,
            error="no archived snapshot",
            via_wayback=True,
        )

    try:
        resp = session.get(snapshot, timeout=timeout, allow_redirects=True)
    except requests.RequestException as exc:
        return DownloadResult(
            ok=False,
            status="wayback_request_exception",
            status_code=None,
            source_url=original_url,
            error=str(exc),
            via_wayback=True,
            wayback_snapshot=snapshot,
        )

    if resp.status_code >= 400:
        return DownloadResult(
            ok=False,
            status="wayback_http_error",
            status_code=resp.status_code,
            source_url=original_url,
            error=f"HTTP {resp.status_code}",
            via_wayback=True,
            wayback_snapshot=snapshot,
        )

    if not resp.content:
        return DownloadResult(
            ok=False,
            status="wayback_empty_content",
            status_code=resp.status_code,
            source_url=original_url,
            error="empty archived response",
            via_wayback=True,
            wayback_snapshot=snapshot,
        )

    if not looks_like_pdf(resp.content, resp.headers.get("content-type", ""), str(resp.url)):
        return DownloadResult(
            ok=False,
            status="wayback_not_pdf",
            status_code=resp.status_code,
            source_url=original_url,
            error=f"content-type={resp.headers.get('content-type', '')}",
            via_wayback=True,
            wayback_snapshot=snapshot,
        )

    out_path.write_bytes(resp.content)
    return DownloadResult(
        ok=True,
        status="downloaded_wayback",
        status_code=resp.status_code,
        source_url=original_url,
        saved_path=out_path,
        via_wayback=True,
        wayback_snapshot=snapshot,
    )


def normalize_columns(df: pd.DataFrame) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for col in df.columns:
        norm = re.sub(r"\s+", "", str(col)).lower()
        mapping[norm] = col
    return mapping


def main() -> None:
    parser = argparse.ArgumentParser(description="Download ICO whitepaper PDFs from ico.xlsx")
    parser.add_argument("--input", default="ico.xlsx", help="Path to input .xlsx file")
    parser.add_argument("--output-dir", default="whitepapers", help="Where PDFs are saved")
    parser.add_argument("--report", default="download_report.csv", help="CSV report path")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="HTTP timeout in seconds")
    parser.add_argument("--sleep", type=float, default=0.2, help="Sleep between requests in seconds")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    report_path = Path(args.report)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(input_path)
    col_map = normalize_columns(df)

    token_col = col_map.get("token")
    link_col = col_map.get("link_white_paper") or col_map.get("link_white_paper\\n")

    if not token_col or not link_col:
        raise ValueError(
            "Cannot find required columns. Need 'token' and 'link_white_paper' (newline tolerated)."
        )

    report_rows = []
    session = requests.Session()
    session.headers.update({"user-agent": "ico-whitepaper-downloader/1.0"})

    for idx, row in df.iterrows():
        token_raw = "" if pd.isna(row[token_col]) else str(row[token_col])
        link_raw = "" if pd.isna(row[link_col]) else str(row[link_col]).strip()

        token_clean = sanitize_filename(token_raw, fallback=f"row_{idx+1}")
        file_path = output_dir / f"{token_clean}.pdf"

        base_info = {
            "row": idx + 1,
            "token": token_raw,
            "filename": file_path.name,
            "original_link": link_raw,
            "saved_path": "",
            "status": "",
            "status_code": "",
            "error": "",
            "download_source": "",
            "wayback_snapshot": "",
        }

        if not link_raw:
            base_info.update(
                {
                    "status": "skipped_empty_link",
                    "error": "link is empty",
                    "download_source": "none",
                }
            )
            report_rows.append(base_info)
            continue

        result = download_pdf(session, link_raw, file_path, args.timeout)
        if result.ok:
            base_info.update(
                {
                    "status": result.status,
                    "status_code": result.status_code,
                    "saved_path": str(result.saved_path),
                    "download_source": "direct",
                }
            )
            report_rows.append(base_info)
            time.sleep(args.sleep)
            continue

        wb_result = download_via_wayback(session, link_raw, file_path, args.timeout)
        if wb_result.ok:
            base_info.update(
                {
                    "status": wb_result.status,
                    "status_code": wb_result.status_code,
                    "saved_path": str(wb_result.saved_path),
                    "download_source": "wayback",
                    "wayback_snapshot": wb_result.wayback_snapshot or "",
                }
            )
            report_rows.append(base_info)
            time.sleep(args.sleep)
            continue

        base_info.update(
            {
                "status": f"failed_direct:{result.status};failed_wayback:{wb_result.status}",
                "status_code": result.status_code or wb_result.status_code or "",
                "error": f"direct={result.error}; wayback={wb_result.error}",
                "download_source": "failed",
                "wayback_snapshot": wb_result.wayback_snapshot or "",
            }
        )
        report_rows.append(base_info)
        time.sleep(args.sleep)

    with report_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(report_rows[0].keys()) if report_rows else [])
        if report_rows:
            writer.writeheader()
            writer.writerows(report_rows)

    total = len(report_rows)
    direct_ok = sum(1 for r in report_rows if r["download_source"] == "direct")
    wb_ok = sum(1 for r in report_rows if r["download_source"] == "wayback")
    failed = sum(1 for r in report_rows if r["download_source"] == "failed")
    skipped = sum(1 for r in report_rows if r["download_source"] == "none")

    print(f"Done. total={total}, direct={direct_ok}, wayback={wb_ok}, failed={failed}, skipped={skipped}")
    print(f"PDF folder: {output_dir.resolve()}")
    print(f"Report: {report_path.resolve()}")


if __name__ == "__main__":
    main()

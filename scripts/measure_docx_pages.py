#!/usr/bin/env python3
"""Measure the real rendered page count of generated DOCX files.

软著材料对页数有硬性要求（操作手册一般不少于 40 页；程序鉴别材料为前 30 页 +
后 30 页、每页不少于 50 行）。仅靠“材料行数 / 每页行数”估算会与 Word 的实际
排版结果不一致（实测偏差常见 1~4 页），交付前必须实测一次。

Windows 上通过 Word COM 的 ComputeStatistics(wdStatisticPages) 取真实页数；
未安装 Word 或非 Windows 环境时回退为按段落行数估算，并在输出中标注口径。

用法：
    python3 measure_docx_pages.py <docx> [<docx> ...]
    python3 measure_docx_pages.py --dir 软件著作权申请资料/正式资料
    python3 measure_docx_pages.py <docx> --manual-min 40 --code-min 60 --json out.json
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

PS_SNIPPET = r"""
$ErrorActionPreference='Continue'
$listFile = "%LIST%"
$outFile = "%OUT%"
$targets = [System.IO.File]::ReadAllLines($listFile, [System.Text.Encoding]::UTF8)
$result = @()
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  foreach ($t in $targets) {
    $p = [string]$t
    if (-not $p) { continue }
    try {
      $doc = $word.Documents.Open($p, $false, $true)
      $doc.Repaginate()
      $result += [PSCustomObject]@{ file=$p; pages=[int]$doc.ComputeStatistics(2); lines=[int]$doc.ComputeStatistics(1); word=$true; error="" }
      $doc.Close($false)
    } catch {
      $result += [PSCustomObject]@{ file=$p; pages=$null; lines=$null; word=$false; error=$_.Exception.Message }
    }
  }
} catch {
  $result += [PSCustomObject]@{ file=""; pages=$null; lines=$null; word=$false; error=$_.Exception.Message }
} finally {
  if ($word) { $word.Quit() }
}
$json = $result | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText($outFile, $json, (New-Object System.Text.UTF8Encoding($false)))
"""


def word_available() -> bool:
    if os.name != "nt":
        return False
    return bool(shutil.which("powershell") or shutil.which("powershell.exe"))


def measure_with_word(paths: list[Path]) -> list[dict] | None:
    """Return [{path, pages, lines}] measured by Word COM, or None if unavailable."""
    if not word_available():
        return None

    tmp_dir = Path(tempfile.mkdtemp(prefix="docx_pages_"))
    list_file = tmp_dir / "targets.txt"
    out_file = tmp_dir / "result.json"
    # 路径写入 UTF-8 清单文件，内联 PowerShell 脚本保持纯 ASCII，
    # 避免中文路径经 .ps1 脚本文件时被编码破坏。
    list_file.write_text("\n".join(str(p.resolve()) for p in paths), encoding="utf-8")
    snippet = PS_SNIPPET.replace("%LIST%", list_file.as_posix()).replace("%OUT%", out_file.as_posix())

    ps = shutil.which("powershell") or shutil.which("powershell.exe")
    try:
        subprocess.run(
            [ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", snippet],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=600,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"[warn] Word 实测失败：{exc}", file=sys.stderr)
        return None

    if not out_file.exists():
        return None
    try:
        data = json.loads(out_file.read_text(encoding="utf-8-sig"))
    except Exception:  # noqa: BLE001
        return None
    if isinstance(data, dict):
        data = [data]
    return [row for row in data if isinstance(row, dict)]


def estimate_pages(path: Path) -> dict:
    """Fallback: count non-empty paragraphs and estimate pages from line count."""
    import re
    import zipfile

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    breaks = len(re.findall(r'w:type="page"', xml))
    paragraphs = len(re.findall(r"<w:p[ >]", xml))
    return {"file": str(path), "pages": breaks + 1, "lines": paragraphs, "word": False,
            "error": "估算口径：按分页符 + 段落数推算，非 Word 实测"}


def rule_check(row: dict, manual_min: int, code_min: int) -> str:
    name = Path(row["file"]).name
    pages = row.get("pages")
    if pages is None:
        return "无法测量"
    if "操作手册" in name or "手册" in name:
        return f"操作手册 {pages} 页（要求 ≥{manual_min}）" + (" ✔" if pages >= manual_min else " ✘")
    if "代码" in name:
        return f"代码材料 {pages} 页（要求 ≥{code_min}）" + (" ✔" if pages >= code_min else " ✘")
    return f"{pages} 页"


def main() -> None:
    parser = argparse.ArgumentParser(description="实测 DOCX 真实页数（Windows 走 Word COM）")
    parser.add_argument("files", nargs="*", help="待测量的 docx 文件")
    parser.add_argument("--dir", help="改为测量该目录下所有 .docx")
    parser.add_argument("--manual-min", type=int, default=40, help="操作手册最少页数，默认 40")
    parser.add_argument("--code-min", type=int, default=60, help="代码材料最少页数，默认 60")
    parser.add_argument("--json", help="把测量结果写入该 JSON 文件")
    args = parser.parse_args()

    targets: list[Path] = []
    if args.dir:
        targets.extend(sorted(Path(args.dir).glob("*.docx")))
    targets.extend(Path(f) for f in args.files)
    targets = [p for p in targets if p.exists()]
    if not targets:
        raise SystemExit("没有可测量的 docx 文件")

    rows = measure_with_word(targets)
    measured_by_word = rows is not None
    if rows is None:
        print("[warn] 未检测到 Word（仅 Windows + Office 可用），改用行数估算口径。", file=sys.stderr)
        rows = [estimate_pages(p) for p in targets]

    print(f"口径：{'Word 实测' if measured_by_word else '行数估算'}")
    failed = False
    for row in rows:
        if not row.get("file"):
            print(f"[warn] {row.get('error')}")
            failed = True
            continue
        verdict = rule_check(row, args.manual_min, args.code_min)
        print(f"  {Path(row['file']).name}: {verdict}")
        if "✘" in verdict or "无法测量" in verdict:
            failed = True

    if args.json:
        Path(args.json).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"结果已写入 {args.json}")

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()

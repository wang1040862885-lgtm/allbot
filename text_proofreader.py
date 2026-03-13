#!/usr/bin/env python3
"""中文文档错别字与病句检测工具（Windows 10 兼容）。"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
import zipfile
import tempfile
from urllib.parse import urlparse
from urllib.request import urlopen
from xml.etree import ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover
    tk = None

try:
    from pycorrector import Corrector
except ImportError:
    Corrector = None


SENTENCE_SPLIT_RE = re.compile(r"[^。！？!?；;\n]+[。！？!?；;]?", re.M)
BAD_SENTENCE_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = (
    ("连续重复标点", re.compile(r"[，。！？；,.!?;]{2,}")),
    ("疑似重复词", re.compile(r"(的的|了了|和和|以及以及|我们我们|他们他们|你你|我我)")),
    ("句子过长且无停顿", re.compile(r"[^。！？!?\n]{120,}")),
    ("关联词疑似搭配不当", re.compile(r"虽然[^。！？!?\n]{0,50}但是[^。！？!?\n]{0,6}却")),
)


@dataclass
class Finding:
    category: str
    original: str
    suggestion: str
    start: int
    end: int
    line: int = 0
    column: int = 0
    context: str = ""


class TextProofreader:
    def __init__(self) -> None:
        self.corrector = Corrector() if Corrector is not None else None

    def split_sentences(self, text: str) -> List[Tuple[str, int]]:
        result: List[Tuple[str, int]] = []
        for m in SENTENCE_SPLIT_RE.finditer(text):
            sent = m.group(0).strip()
            if sent:
                result.append((sent, m.start()))
        return result

    def detect_typos(self, text: str) -> List[Finding]:
        findings: List[Finding] = []
        if self.corrector is None:
            return findings
        for sentence, offset in self.split_sentences(text):
            _, details = self.corrector.correct(sentence)
            for detail in details:
                # pycorrector detail: [error, correct, begin_idx, end_idx]
                wrong, right, begin_idx, end_idx = detail[:4]
                findings.append(
                    Finding(
                        category="错别字",
                        original=wrong,
                        suggestion=right,
                        start=offset + begin_idx,
                        end=offset + end_idx,
                    )
                )
        return findings

    def detect_bad_sentences(self, text: str) -> List[Finding]:
        findings: List[Finding] = []
        for reason, pattern in BAD_SENTENCE_PATTERNS:
            for m in pattern.finditer(text):
                findings.append(
                    Finding(
                        category=f"病句({reason})",
                        original=m.group(0),
                        suggestion="建议人工复核上下文并优化表达",
                        start=m.start(),
                        end=m.end(),
                    )
                )
        return findings

    def scan(self, text: str) -> List[Finding]:
        findings = self.detect_typos(text)
        findings.extend(self.detect_bad_sentences(text))
        self.enrich_positions(text, findings)
        findings.sort(key=lambda f: f.start)
        return findings

    @staticmethod
    def enrich_positions(text: str, findings: List[Finding], context_size: int = 16) -> None:
        """补充 line/column 以及上下文片段。"""
        line_starts = [0]
        for m in re.finditer(r"\n", text):
            line_starts.append(m.end())

        import bisect

        for item in findings:
            line_idx = bisect.bisect_right(line_starts, item.start) - 1
            item.line = line_idx + 1
            item.column = item.start - line_starts[line_idx] + 1
            left = max(0, item.start - context_size)
            right = min(len(text), item.end + context_size)
            item.context = text[left:right].replace("\n", " ")


def read_text(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    if ext in {".txt", ".md", ".csv", ".log"}:
        return file_path.read_text(encoding="utf-8", errors="ignore")
    if ext == ".docx":
        return read_docx_text(file_path)
    if ext == ".doc":
        raise ValueError("暂不支持 .doc（二进制旧格式），请先另存为 .docx")
    raise ValueError(f"不支持的文件类型: {ext}")


def read_docx_text(file_path: Path) -> str:
    """轻量读取 docx 文本，无需额外依赖（适配离线环境）。"""
    try:
        with zipfile.ZipFile(file_path) as zf:
            xml_bytes = zf.read("word/document.xml")
    except KeyError as exc:
        raise RuntimeError(f"无效的 docx 文件: 缺少 document.xml ({file_path})") from exc

    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    paragraphs: List[str] = []
    for para in root.findall(".//w:p", ns):
        texts = [node.text or "" for node in para.findall('.//w:t', ns)]
        paragraphs.append("".join(texts))
    return "\n".join(paragraphs)



def find_docx_files(search_dir: Path) -> List[Path]:
    """递归查找目录中的 docx 文件（排除临时文件）。"""
    files = [
        p for p in search_dir.rglob('*.docx')
        if p.is_file() and not p.name.startswith('~$')
    ]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files



def normalize_input_path(raw_path: str) -> Path:
    """清理输入路径中的换行和首尾引号。"""
    cleaned = raw_path.strip().strip('"').strip("'").replace("\r", "").replace("\n", "")
    return Path(cleaned)


def to_github_raw_url(url: str) -> str:
    """将 GitHub 页面 URL 转成 raw 下载 URL。"""
    if "raw.githubusercontent.com" in url:
        return url

    parsed = urlparse(url)
    if parsed.netloc not in {"github.com", "www.github.com"}:
        raise ValueError("仅支持 github.com / raw.githubusercontent.com 链接")

    parts = [p for p in parsed.path.split("/") if p]
    # owner/repo/blob/branch/path/to/file.docx
    if len(parts) >= 5 and parts[2] == "blob":
        owner, repo, _, branch, *path_parts = parts
        return f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{'/'.join(path_parts)}"

    raise ValueError("无法识别 GitHub 文件链接，请使用 blob 链接或 raw 链接")


def download_github_docx(url: str) -> Path:
    """下载 GitHub 上的 docx 到临时文件并返回本地路径。"""
    raw_url = to_github_raw_url(url)
    with urlopen(raw_url, timeout=30) as resp:
        data = resp.read()

    fd, tmp_name = tempfile.mkstemp(prefix="docx_scan_", suffix=".docx")
    Path(tmp_name).write_bytes(data)
    os.close(fd)
    return Path(tmp_name)



def write_report(findings: Iterable[Finding], output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8") as f:
        f.write("category\toriginal\tsuggestion\tstart\tend\tline\tcolumn\tcontext\n")
        for item in findings:
            f.write(
                f"{item.category}\t{item.original}\t{item.suggestion}\t{item.start}\t{item.end}"
                f"\t{item.line}\t{item.column}\t{item.context}\n"
            )


def run_cli(file_path: Path, output_path: Path | None) -> int:
    text = read_text(file_path)
    scanner = TextProofreader()

    t0 = time.perf_counter()
    findings = scanner.scan(text)
    duration = time.perf_counter() - t0

    print(f"扫描完成，总字符数: {len(text):,}")
    print(f"发现问题数: {len(findings)}")
    print(f"耗时: {duration:.2f}s")
    if Corrector is None:
        print("提示: 未安装 pycorrector，本次仅执行病句规则检测。")

    preview = findings[:30]
    for idx, item in enumerate(preview, 1):
        print(
            f"[{idx:02d}] {item.category} | '{item.original}' -> '{item.suggestion}' | "
            f"位置: {item.start}-{item.end} (行{item.line}, 列{item.column}) | 上下文: {item.context}"
        )

    if output_path:
        write_report(findings, output_path)
        print(f"已输出报告: {output_path}")

    return 0


class ProofreaderGUI:
    def __init__(self) -> None:
        if tk is None:
            raise RuntimeError("当前环境不可用 Tkinter GUI")
        self.root = tk.Tk()
        self.root.title("文档错别字与病句扫描器")
        self.root.geometry("1080x680")

        top = ttk.Frame(self.root)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(top, text="选择文件并扫描", command=self.choose_and_scan).pack(side=tk.LEFT)
        self.status = ttk.Label(top, text="就绪")
        self.status.pack(side=tk.LEFT, padx=12)

        cols = ("category", "original", "suggestion", "start", "end")
        self.table = ttk.Treeview(self.root, columns=cols, show="headings")
        for c, w in zip(cols, (150, 220, 350, 120, 120)):
            self.table.heading(c, text=c)
            self.table.column(c, width=w, anchor=tk.W)
        self.table.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    def choose_and_scan(self) -> None:
        path = filedialog.askopenfilename(
            title="请选择要扫描的文档",
            filetypes=[
                ("支持文件", "*.txt *.md *.docx *.csv *.log"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        try:
            text = read_text(Path(path))
            scanner = TextProofreader()
            t0 = time.perf_counter()
            findings = scanner.scan(text)
            cost = time.perf_counter() - t0

            for row in self.table.get_children():
                self.table.delete(row)
            for item in findings:
                self.table.insert(
                    "",
                    tk.END,
                    values=(
                        item.category,
                        item.original,
                        item.suggestion,
                        f"{item.start} (L{item.line}:C{item.column})",
                        item.end,
                    ),
                )

            self.status.configure(
                text=f"扫描完成：{len(text):,} 字符，{len(findings)} 个问题，耗时 {cost:.2f}s"
            )
            messagebox.showinfo(
                "扫描完成",
                f"文件: {path}\n总字符: {len(text):,}\n问题数: {len(findings)}\n耗时: {cost:.2f}s",
            )
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc))
            self.status.configure(text=f"失败: {exc}")

    def run(self) -> None:
        self.root.mainloop()


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="中文文档错别字和病句检测工具")
    p.add_argument("file", nargs="?", help="目标文件路径（txt/md/docx/csv/log）")
    p.add_argument("-o", "--output", help="输出报告路径（tsv）")
    p.add_argument("--gui", action="store_true", help="启动图形界面")
    p.add_argument("--auto-docx", action="store_true", help="自动查找当前目录下最近修改的 docx 并扫描")
    p.add_argument("--search-dir", default=".", help="自动查找 docx 时的搜索目录，默认当前目录")
    p.add_argument("--github-url", help="直接传入 GitHub 上 docx 文件链接（blob 或 raw）")
    p.add_argument("--list-docx", action="store_true", help="列出搜索目录下检测到的 docx 路径")
    return p


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.gui:
        ProofreaderGUI().run()
        return 0

    file_path: Path | None = None
    temp_download: Path | None = None

    if args.list_docx:
        search_dir = Path(args.search_dir)
        if not search_dir.exists() or not search_dir.is_dir():
            print(f"搜索目录无效: {search_dir}", file=sys.stderr)
            return 2
        files = find_docx_files(search_dir)
        if not files:
            print(f"未找到 docx: {search_dir.resolve()}")
            return 0
        print("找到以下 docx 文件：")
        for idx, path in enumerate(files, 1):
            print(f"[{idx:02d}] {path}")
        return 0

    if args.github_url:
        try:
            temp_download = download_github_docx(args.github_url)
        except Exception as exc:
            print(f"下载 GitHub docx 失败: {exc}", file=sys.stderr)
            return 2
        file_path = temp_download
        print(f"已下载 GitHub docx 到临时文件: {file_path}")
    elif args.file:
        file_path = normalize_input_path(args.file)
        if not file_path.exists():
            print(f"文件不存在: {file_path}", file=sys.stderr)
            return 2
    elif args.auto_docx:
        search_dir = Path(args.search_dir)
        if not search_dir.exists() or not search_dir.is_dir():
            print(f"搜索目录无效: {search_dir}", file=sys.stderr)
            return 2
        docx_files = find_docx_files(search_dir)
        if not docx_files:
            print(f"未在目录中找到 docx 文件: {search_dir.resolve()}", file=sys.stderr)
            return 2
        file_path = docx_files[0]
        print(f"自动选择 docx: {file_path}")
    else:
        print("请指定目标文件路径，或使用 --auto-docx 自动查找，或 --github-url 下载并扫描，或 --gui 启动图形界面", file=sys.stderr)
        return 2

    output_path = Path(args.output) if args.output else None
    try:
        return run_cli(file_path, output_path)
    finally:
        if temp_download and temp_download.exists():
            temp_download.unlink(missing_ok=True)


if __name__ == "__main__":
    raise SystemExit(main())

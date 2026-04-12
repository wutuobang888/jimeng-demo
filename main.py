from __future__ import annotations

import os
import re
import shutil
import tempfile
import traceback
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
LOW_CONFIDENCE_THRESHOLD = 60.0
DEFAULT_OUTPUT_FOLDER_NAME = "sorted_output"


pd: Any | None = None
fuzz: Any | None = None
linear_sum_assignment: Any | None = None


def load_runtime_dependencies() -> None:
    global pd, fuzz, linear_sum_assignment
    if pd is None:
        import pandas as _pd

        pd = _pd
    if fuzz is None:
        from rapidfuzz import fuzz as _fuzz

        fuzz = _fuzz
    if linear_sum_assignment is None:
        from scipy.optimize import linear_sum_assignment as _lsa

        linear_sum_assignment = _lsa


@dataclass
class SourceItem:
    source_type: str  # folder | zip
    path: Path

    @property
    def key(self) -> str:
        return f"{self.source_type}|{str(self.path.resolve()).lower()}"

    @property
    def label(self) -> str:
        return str(self.path)


@dataclass
class ImageRecord:
    path: Path
    source_path: Path
    source_type: str
    source_rel_path: str
    original_name: str
    extracted_text: str


@dataclass
class MatchResult:
    prompt_index: int
    prompt_text: str
    image: ImageRecord
    score: float
    low_confidence: bool


@dataclass
class SourceScanResult:
    images: list[ImageRecord]
    errors: list[str]


def parse_prompts(raw_text: str) -> list[str]:
    prompts = [line.strip() for line in raw_text.splitlines() if line.strip()]
    return prompts


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def extract_semantic_text(filename: str) -> str:
    stem = Path(filename).stem
    text = stem

    # 常见平台前缀和噪声字段处理
    text = re.sub(r"(?i)\b(jimeng|即梦|jianying|doubao|wechat|weixin|img|image|photo)\b", " ", text)
    text = re.sub(r"\d{4}[-_]?\d{1,2}[-_]?\d{1,2}", " ", text)  # 日期
    text = re.sub(r"\d{6,}", " ", text)  # 长数字编号
    text = re.sub(r"[_\-|]+", " ", text)
    text = re.sub(r"[()\[\]{}]+", " ", text)
    text = re.sub(r"[，。；：、‘’“”！？,.:;!?]+", " ", text)

    tokens = [t.strip() for t in text.split() if t.strip()]
    filtered: list[str] = []
    for token in tokens:
        low = token.lower()
        if low in {"copy", "final", "draft"}:
            continue
        if re.fullmatch(r"[a-f0-9]{8,}", low):
            continue
        if re.fullmatch(r"\d{1,4}", token):
            continue
        filtered.append(token)

    cleaned = " ".join(filtered)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    # 尽量保留中英文和数字语义
    chunks = re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", cleaned)
    if chunks:
        cleaned = " ".join(chunks)

    return normalize_whitespace(cleaned)


def normalize_for_match(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[，。；：、‘’“”！？,.:;!?\"'()\[\]{}<>/\\|_`~@#$%^&*+=-]+", " ", text)
    return normalize_whitespace(text)


def char_token_text(text: str) -> str:
    compact = normalize_for_match(text).replace(" ", "")
    return " ".join(compact)


def compute_match_score(prompt: str, image_text: str) -> float:
    if fuzz is None:
        raise RuntimeError("Runtime dependencies are not loaded")

    prompt_norm = normalize_for_match(prompt)
    image_norm = normalize_for_match(image_text)
    if not image_norm:
        return 0.0

    prompt_chars = char_token_text(prompt_norm)
    image_chars = char_token_text(image_norm)

    s1 = fuzz.partial_ratio(prompt_norm, image_norm)
    s2 = fuzz.token_sort_ratio(prompt_chars, image_chars)
    s3 = fuzz.token_set_ratio(prompt_chars, image_chars)

    score = 0.45 * s1 + 0.30 * s2 + 0.25 * s3
    return round(score, 2)


def make_unique_output_dir(base_dir: Path, folder_name: str = DEFAULT_OUTPUT_FOLDER_NAME) -> Path:
    target = base_dir / folder_name
    if not target.exists():
        target.mkdir(parents=True, exist_ok=False)
        return target

    idx = 1
    while True:
        candidate = base_dir / f"{folder_name}_{idx}"
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
        idx += 1


def safe_filename_fragment(text: str, max_length: int = 24) -> str:
    text = normalize_whitespace(text)
    text = re.sub(r'[<>:"/\\|?*]', "", text)
    text = text.replace("\n", " ").replace("\r", " ").strip(" .")
    if not text:
        return "untitled"
    return text[:max_length].strip() or "untitled"


def scan_folder_source(source: SourceItem) -> SourceScanResult:
    images: list[ImageRecord] = []
    errors: list[str] = []

    if not source.path.exists():
        return SourceScanResult(images=[], errors=[f"来源不存在：{source.path}"])
    if not source.path.is_dir():
        return SourceScanResult(images=[], errors=[f"不是有效文件夹：{source.path}"])

    def _on_walk_error(exc: OSError) -> None:
        errors.append(f"无权限或读取失败：{exc}")

    for root, _, files in os.walk(source.path, onerror=_on_walk_error):
        root_path = Path(root)
        for name in files:
            full = root_path / name
            if full.suffix.lower() not in IMAGE_EXTENSIONS:
                continue
            try:
                rel = str(full.relative_to(source.path))
            except ValueError:
                rel = name
            images.append(
                ImageRecord(
                    path=full,
                    source_path=source.path,
                    source_type="文件夹",
                    source_rel_path=rel,
                    original_name=full.name,
                    extracted_text=extract_semantic_text(full.name),
                )
            )

    return SourceScanResult(images=images, errors=errors)


def scan_zip_source(source: SourceItem, temp_root: Path) -> SourceScanResult:
    images: list[ImageRecord] = []
    errors: list[str] = []

    if not source.path.exists():
        return SourceScanResult(images=[], errors=[f"来源不存在：{source.path}"])
    if not source.path.is_file():
        return SourceScanResult(images=[], errors=[f"不是有效压缩包文件：{source.path}"])
    if source.path.suffix.lower() != ".zip":
        return SourceScanResult(images=[], errors=[f"暂不支持该压缩格式：{source.path.name}"])

    extract_dir = temp_root / f"zip_{source.path.stem}"
    extract_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(source.path, "r") as zf:
            zf.extractall(extract_dir)
    except zipfile.BadZipFile:
        return SourceScanResult(images=[], errors=[f"压缩包损坏或无法解压：{source.path}"])
    except Exception as exc:  # noqa: BLE001
        return SourceScanResult(images=[], errors=[f"解压失败：{source.path}，原因：{exc}"])

    for root, _, files in os.walk(extract_dir):
        root_path = Path(root)
        for name in files:
            full = root_path / name
            if full.suffix.lower() not in IMAGE_EXTENSIONS:
                continue
            try:
                rel = str(full.relative_to(extract_dir))
            except ValueError:
                rel = name
            images.append(
                ImageRecord(
                    path=full,
                    source_path=source.path,
                    source_type="压缩包",
                    source_rel_path=rel,
                    original_name=full.name,
                    extracted_text=extract_semantic_text(full.name),
                )
            )

    return SourceScanResult(images=images, errors=errors)


def build_matches(prompts: list[str], images: list[ImageRecord]) -> tuple[list[MatchResult], list[int], list[int]]:
    if linear_sum_assignment is None:
        raise RuntimeError("Runtime dependencies are not loaded")

    if not prompts or not images:
        return [], list(range(len(prompts))), list(range(len(images)))

    score_matrix: list[list[float]] = []
    for prompt in prompts:
        row = [compute_match_score(prompt, img.extracted_text) for img in images]
        score_matrix.append(row)

    max_score = 100.0
    cost_matrix = [[max_score - s for s in row] for row in score_matrix]
    p_idx, i_idx = linear_sum_assignment(cost_matrix)

    matches: list[MatchResult] = []
    used_p: set[int] = set()
    used_i: set[int] = set()

    for a, b in zip(p_idx.tolist(), i_idx.tolist()):
        score = score_matrix[a][b]
        img = images[b]
        matches.append(
            MatchResult(
                prompt_index=a + 1,
                prompt_text=prompts[a],
                image=img,
                score=score,
                low_confidence=(score < LOW_CONFIDENCE_THRESHOLD or not img.extracted_text),
            )
        )
        used_p.add(a)
        used_i.add(b)

    matches.sort(key=lambda x: x.prompt_index)
    unmatched_prompts = [i for i in range(len(prompts)) if i not in used_p]
    unmatched_images = [i for i in range(len(images)) if i not in used_i]
    return matches, unmatched_prompts, unmatched_images


def export_results(
    matches: list[MatchResult],
    prompts: list[str],
    images: list[ImageRecord],
    unmatched_prompts: list[int],
    unmatched_images: list[int],
    output_dir: Path,
) -> tuple[Any, int]:
    if pd is None:
        raise RuntimeError("Runtime dependencies are not loaded")

    rows: list[dict[str, Any]] = []
    low_count = 0

    for m in matches:
        idx_str = f"{m.prompt_index:03d}"
        ext = m.image.path.suffix.lower()
        frag = safe_filename_fragment(m.prompt_text)
        export_name = f"{idx_str}_{frag}{ext}"
        export_path = output_dir / export_name

        status = "已匹配"
        err_msg = ""
        try:
            shutil.copy2(m.image.path, export_path)
        except Exception as exc:  # noqa: BLE001
            status = "导出失败"
            err_msg = str(exc)

        if m.low_confidence:
            low_count += 1

        rows.append(
            {
                "目标序号": m.prompt_index,
                "目标提示词": m.prompt_text,
                "原文件名": m.image.original_name,
                "文件名提取文本": m.image.extracted_text,
                "匹配得分": m.score,
                "是否低置信度": "是" if m.low_confidence else "否",
                "导出文件名": export_name if status != "导出失败" else "",
                "来源路径": str(m.image.source_path),
                "来源类型": m.image.source_type,
                "原始相对路径": m.image.source_rel_path,
                "状态": status,
                "错误信息": err_msg,
            }
        )

    for p_idx in unmatched_prompts:
        rows.append(
            {
                "目标序号": p_idx + 1,
                "目标提示词": prompts[p_idx],
                "原文件名": "",
                "文件名提取文本": "",
                "匹配得分": 0,
                "是否低置信度": "是",
                "导出文件名": "",
                "来源路径": "",
                "来源类型": "",
                "原始相对路径": "",
                "状态": "未分配到图片",
                "错误信息": "图片数量不足或匹配失败",
            }
        )

    for i_idx in unmatched_images:
        img = images[i_idx]
        rows.append(
            {
                "目标序号": "",
                "目标提示词": "",
                "原文件名": img.original_name,
                "文件名提取文本": img.extracted_text,
                "匹配得分": "",
                "是否低置信度": "",
                "导出文件名": "",
                "来源路径": str(img.source_path),
                "来源类型": img.source_type,
                "原始相对路径": img.source_rel_path,
                "状态": "图片未被分配",
                "错误信息": "提示词数量不足或全局分配未命中",
            }
        )

    rows.sort(
        key=lambda x: (
            x["目标序号"] if isinstance(x["目标序号"], int) else 10**9,
            str(x["原文件名"]),
        )
    )

    df = pd.DataFrame(rows)
    df.to_csv(output_dir / "match_results.csv", index=False, encoding="utf-8-sig")
    return df, low_count


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("即梦图片自动匹配排序工具")
        self.root.geometry("1120x820")
        self.root.minsize(980, 720)

        self.sources: list[SourceItem] = []

        self.output_dir_var = tk.StringVar(value="未选择（默认输出到程序目录）")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

        title = ttk.Label(frame, text="即梦图片自动匹配排序工具（MVP）", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w")

        # 提示词输入区
        prompt_box = ttk.LabelFrame(frame, text="提示词输入区（一行一条，顺序即目标顺序）")
        prompt_box.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        prompt_box.columnconfigure(0, weight=1)
        prompt_box.rowconfigure(0, weight=1)

        self.prompt_text = scrolledtext.ScrolledText(prompt_box, wrap=tk.WORD, height=12, font=("Microsoft YaHei UI", 10))
        self.prompt_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        # 来源管理区
        source_box = ttk.LabelFrame(frame, text="图片来源（支持文件夹 + ZIP 压缩包混合）")
        source_box.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        source_box.columnconfigure(0, weight=1)
        source_box.rowconfigure(1, weight=1)

        button_row = ttk.Frame(source_box)
        button_row.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        ttk.Button(button_row, text="添加文件夹来源", command=self.add_folder_source).pack(side=tk.LEFT)
        ttk.Button(button_row, text="添加 ZIP 来源（可多选）", command=self.add_zip_sources).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(button_row, text="移除选中来源", command=self.remove_selected_source).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(button_row, text="清空全部来源", command=self.clear_sources).pack(side=tk.LEFT, padx=(8, 0))

        columns = ("type", "path")
        self.source_tree = ttk.Treeview(source_box, columns=columns, show="headings", height=8)
        self.source_tree.heading("type", text="来源类型")
        self.source_tree.heading("path", text="来源路径")
        self.source_tree.column("type", width=90, anchor="center")
        self.source_tree.column("path", width=880, anchor="w")
        self.source_tree.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        # 输出目录
        out_box = ttk.LabelFrame(frame, text="输出目录")
        out_box.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        out_box.columnconfigure(0, weight=1)

        ttk.Label(out_box, textvariable=self.output_dir_var).grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        ttk.Button(out_box, text="选择输出目录", command=self.select_output_dir).grid(row=0, column=1, padx=8, pady=8)

        # 操作按钮
        action_row = ttk.Frame(frame)
        action_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))

        self.start_btn = ttk.Button(action_row, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT)

        # 日志区
        log_box = ttk.LabelFrame(frame, text="日志")
        log_box.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_box, wrap=tk.WORD, height=12, state=tk.DISABLED, font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

    def log(self, msg: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()

    def add_source(self, source: SourceItem) -> None:
        for s in self.sources:
            if s.key == source.key:
                self.log(f"来源已存在，已忽略：{source.path}")
                messagebox.showinfo("重复来源", f"该来源已存在，已忽略：\n{source.path}")
                return
        self.sources.append(source)
        display_type = "文件夹" if source.source_type == "folder" else "压缩包"
        self.source_tree.insert("", tk.END, values=(display_type, str(source.path)))
        self.log(f"已添加来源：[{display_type}] {source.path}")

    def add_folder_source(self) -> None:
        folder = filedialog.askdirectory(title="选择图片文件夹来源")
        if not folder:
            return
        self.add_source(SourceItem(source_type="folder", path=Path(folder)))

    def add_zip_sources(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 ZIP 压缩包来源（可多选）",
            filetypes=[("ZIP 文件", "*.zip"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            self.add_source(SourceItem(source_type="zip", path=Path(p)))

    def remove_selected_source(self) -> None:
        selected = self.source_tree.selection()
        if not selected:
            messagebox.showwarning("未选择", "请先在来源列表中选择要移除的来源")
            return

        remove_keys: set[str] = set()
        for item_id in selected:
            vals = self.source_tree.item(item_id, "values")
            src_type_display, src_path = vals[0], vals[1]
            src_type = "folder" if src_type_display == "文件夹" else "zip"
            key = f"{src_type}|{str(Path(src_path).resolve()).lower()}"
            remove_keys.add(key)
            self.source_tree.delete(item_id)

        self.sources = [s for s in self.sources if s.key not in remove_keys]
        self.log(f"已移除来源 {len(remove_keys)} 项")

    def clear_sources(self) -> None:
        if not self.sources:
            return
        if not messagebox.askyesno("确认清空", "确认清空全部来源吗？"):
            return
        self.sources.clear()
        for iid in self.source_tree.get_children():
            self.source_tree.delete(iid)
        self.log("已清空全部来源")

    def select_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir_var.set(folder)
            self.log(f"已选择输出目录：{folder}")

    def get_output_base(self) -> Path:
        val = self.output_dir_var.get().strip()
        if not val or val == "未选择（默认输出到程序目录）":
            return Path(__file__).resolve().parent
        p = Path(val)
        if not p.exists():
            raise FileNotFoundError("所选输出目录不存在")
        if not p.is_dir():
            raise NotADirectoryError("所选输出路径不是文件夹")
        return p

    def run_pipeline(self) -> Path:
        prompts = parse_prompts(self.prompt_text.get("1.0", tk.END))
        if not prompts:
            raise ValueError("请输入提示词内容")
        if not self.sources:
            raise ValueError("请先添加图片来源")

        self.log(f"提示词条数：{len(prompts)}")
        self.log(f"来源条数：{len(self.sources)}")
        self.log("正在加载依赖（rapidfuzz / scipy / pandas）...")
        load_runtime_dependencies()
        self.log("依赖加载完成")

        all_images: list[ImageRecord] = []
        scan_errors: list[str] = []

        with tempfile.TemporaryDirectory(prefix="jimeng_matcher_") as tmp:
            tmp_root = Path(tmp)

            for idx, src in enumerate(self.sources, start=1):
                self.log(f"[{idx}/{len(self.sources)}] 扫描来源：{src.path}")
                if src.source_type == "folder":
                    result = scan_folder_source(src)
                else:
                    result = scan_zip_source(src, tmp_root)

                if result.errors:
                    for err in result.errors:
                        self.log(f"来源异常：{err}")
                        scan_errors.append(err)

                if not result.images:
                    self.log(f"该来源未发现可用图片：{src.path}")
                else:
                    self.log(f"该来源扫描到图片：{len(result.images)} 张")

                all_images.extend(result.images)

            if not all_images:
                raise FileNotFoundError("所有来源中都未发现 png/jpg/jpeg 图片")

            self.log(f"总图片数：{len(all_images)}")
            empty_extract = sum(1 for i in all_images if not i.extracted_text)
            if empty_extract:
                self.log(f"文件名提取后为空：{empty_extract} 张（将标记低置信度）")

            if len(prompts) != len(all_images):
                self.log(f"提示词数量与图片数量不一致：提示词 {len(prompts)}，图片 {len(all_images)}")

            self.log("正在执行全局唯一匹配...")
            matches, unmatched_prompts, unmatched_images = build_matches(prompts, all_images)
            self.log(f"匹配完成：{len(matches)} 条")

            output_base = self.get_output_base()
            out_dir = make_unique_output_dir(output_base)
            self.log(f"输出目录：{out_dir}")

            df, low_count = export_results(
                matches=matches,
                prompts=prompts,
                images=all_images,
                unmatched_prompts=unmatched_prompts,
                unmatched_images=unmatched_images,
                output_dir=out_dir,
            )

            export_fail_count = int((df["状态"] == "导出失败").sum())
            self.log(f"低置信度条数：{low_count}")
            self.log(f"未分配提示词：{len(unmatched_prompts)}")
            self.log(f"未分配图片：{len(unmatched_images)}")
            if export_fail_count:
                self.log(f"导出失败数量：{export_fail_count}")
            if scan_errors:
                self.log(f"来源异常数量：{len(scan_errors)}（详情见上方日志）")

            self.log(f"CSV 已生成：{out_dir / 'match_results.csv'}")
            self.log("处理完成")
            return out_dir

    def start_processing(self) -> None:
        self.start_btn.configure(state=tk.DISABLED)
        self.log("=" * 44)
        self.log("开始处理")
        try:
            out_dir = self.run_pipeline()
        except Exception as exc:  # noqa: BLE001
            self.log(f"错误：{exc}")
            self.log(traceback.format_exc())
            messagebox.showerror("处理失败", str(exc))
        else:
            messagebox.showinfo("处理完成", f"处理完成，输出目录：\n{out_dir}")
        finally:
            self.start_btn.configure(state=tk.NORMAL)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    app = App(root)
    app.log("程序已启动，请粘贴提示词并添加来源")
    root.mainloop()


if __name__ == "__main__":
    main()

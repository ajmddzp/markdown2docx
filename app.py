from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pypandoc

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


TAB_MARKDOWN_TO_WORD = 0
TAB_FILE_TO_MARKDOWN = 1


def normalize_filename(filename: str, extension: str) -> str:
    name = filename.strip()
    if not name:
        raise ValueError("文件名不能为空。")

    ext = extension if extension.startswith(".") else f".{extension}"
    if not name.lower().endswith(ext.lower()):
        name += ext
    return name


def show_warning(title: str, message: str) -> None:
    messagebox.showwarning(title, message)


def show_error(title: str, message: str) -> None:
    messagebox.showerror(title, message)


def choose_output_folder(folder_var: tk.StringVar) -> None:
    selected = filedialog.askdirectory(title="选择输出目录")
    if selected:
        folder_var.set(selected)


def choose_input_file(input_var: tk.StringVar) -> None:
    selected = filedialog.askopenfilename(
        title="选择 Word 或 PDF 文件",
        filetypes=[
            ("Word/PDF 文件", "*.docx *.pdf"),
            ("Word 文件", "*.docx"),
            ("PDF 文件", "*.pdf"),
            ("所有文件", "*.*"),
        ],
    )
    if selected:
        input_var.set(selected)


def convert_markdown_to_docx(markdown_text: str, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    pypandoc.convert_text(
        source=markdown_text,
        to="docx",
        format="md",
        outputfile=str(output_file),
    )


def convert_docx_to_markdown(input_file: Path) -> str:
    try:
        return pypandoc.convert_file(
            source_file=str(input_file),
            to="gfm",
            format="docx",
        )
    except OSError as exc:
        raise RuntimeError(
            "未找到 Pandoc。请先安装 Pandoc。\n"
            "安装地址: https://pandoc.org/installing.html"
        ) from exc
    except RuntimeError as exc:
        raise RuntimeError(f"Word 转 Markdown 失败：{exc}") from exc


def _normalize_pdf_text(raw_text: str) -> str:
    lines = [line.strip() for line in raw_text.splitlines()]
    non_empty = [line for line in lines if line]
    return "\n".join(non_empty)


def convert_pdf_to_markdown(input_file: Path) -> str:
    if PdfReader is None:
        raise RuntimeError("缺少 pypdf 依赖，请先安装：pip install pypdf")

    try:
        reader = PdfReader(str(input_file))
    except Exception as exc:
        raise RuntimeError(f"无法读取 PDF：{exc}") from exc

    page_blocks: list[str] = []
    for page_no, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        cleaned = _normalize_pdf_text(text)
        if cleaned:
            page_blocks.append(f"## 第 {page_no} 页\n\n{cleaned}")

    if not page_blocks:
        raise RuntimeError("PDF 中未提取到可用文本，可能是扫描版图片 PDF。")

    return "\n\n".join(page_blocks)


def write_markdown(markdown_text: str, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(markdown_text, encoding="utf-8")


def export_to_word(
    markdown_input: tk.Text,
    folder_var: tk.StringVar,
    filename_var: tk.StringVar,
) -> None:
    markdown_text = markdown_input.get("1.0", tk.END).strip()
    if not markdown_text:
        show_warning("缺少内容", "请先输入 Markdown 内容。")
        return

    folder_text = folder_var.get().strip()
    if not folder_text:
        show_warning("缺少目录", "请选择输出目录。")
        return

    try:
        output_name = normalize_filename(filename_var.get(), ".docx")
    except ValueError as exc:
        show_warning("文件名无效", str(exc))
        return

    output_path = Path(folder_text) / output_name

    try:
        convert_markdown_to_docx(markdown_text, output_path)
    except OSError:
        show_error(
            "Pandoc 未安装",
            "未找到 Pandoc，请先安装后再转换。\n\nhttps://pandoc.org/installing.html",
        )
        return
    except RuntimeError as exc:
        show_error("转换失败", f"导出 Word 失败。\n\n{exc}")
        return

    messagebox.showinfo("完成", f"Word 文件已导出：\n{output_path}")


def export_to_markdown(
    input_var: tk.StringVar,
    folder_var: tk.StringVar,
    filename_var: tk.StringVar,
) -> None:
    input_text = input_var.get().strip()
    if not input_text:
        show_warning("缺少输入文件", "请先选择 Word 或 PDF 文件。")
        return

    input_file = Path(input_text)
    if not input_file.exists():
        show_error("文件不存在", f"找不到输入文件：\n{input_file}")
        return

    output_folder_text = folder_var.get().strip()
    if not output_folder_text:
        show_warning("缺少保存目录", "请选择 Markdown 保存目录。")
        return

    try:
        output_name = normalize_filename(filename_var.get(), ".md")
    except ValueError as exc:
        show_warning("文件名无效", str(exc))
        return

    output_path = Path(output_folder_text) / output_name
    suffix = input_file.suffix.lower()

    try:
        if suffix == ".docx":
            markdown = convert_docx_to_markdown(input_file)
        elif suffix == ".pdf":
            markdown = convert_pdf_to_markdown(input_file)
        else:
            show_warning("不支持的格式", "目前仅支持 .docx 和 .pdf 文件。")
            return

        write_markdown(markdown, output_path)
    except RuntimeError as exc:
        show_error("转换失败", str(exc))
        return
    except OSError as exc:
        show_error("保存失败", f"无法保存 Markdown 文件。\n\n{exc}")
        return

    messagebox.showinfo("完成", f"Markdown 文件已保存：\n{output_path}")


def build_markdown_to_word_tab(parent: ttk.Frame) -> None:
    ttk.Label(parent, text="Markdown -> Word (.docx)").pack(anchor=tk.W, pady=(0, 8))

    editor_frame = ttk.Frame(parent)
    editor_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    markdown_input = tk.Text(editor_frame, wrap=tk.WORD, undo=True)
    scroll = ttk.Scrollbar(editor_frame, orient=tk.VERTICAL, command=markdown_input.yview)
    markdown_input.configure(yscrollcommand=scroll.set)
    markdown_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    folder_var = tk.StringVar(value=str(Path.cwd()))
    filename_var = tk.StringVar(value="output.docx")

    output_frame = ttk.Frame(parent)
    output_frame.pack(fill=tk.X, pady=(0, 8))

    ttk.Label(output_frame, text="输出目录").grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Entry(output_frame, textvariable=folder_var).grid(row=0, column=1, sticky="ew")
    ttk.Button(
        output_frame,
        text="选择目录",
        command=lambda: choose_output_folder(folder_var),
    ).grid(row=0, column=2, padx=(8, 0))

    ttk.Label(output_frame, text="文件名").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
    ttk.Entry(output_frame, textvariable=filename_var).grid(row=1, column=1, sticky="ew", pady=(8, 0))

    output_frame.columnconfigure(1, weight=1)

    ttk.Button(
        parent,
        text="导出 Word",
        command=lambda: export_to_word(markdown_input, folder_var, filename_var),
    ).pack(anchor=tk.E)


def build_file_to_markdown_tab(parent: ttk.Frame) -> None:
    ttk.Label(parent, text="Word/PDF -> Markdown (.md)").pack(anchor=tk.W, pady=(0, 10))

    input_var = tk.StringVar()
    folder_var = tk.StringVar(value=str(Path.cwd()))
    filename_var = tk.StringVar(value="output.md")

    form = ttk.Frame(parent)
    form.pack(fill=tk.X)

    ttk.Label(form, text="输入文件").grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Entry(form, textvariable=input_var).grid(row=0, column=1, sticky="ew")
    ttk.Button(
        form,
        text="选择文件",
        command=lambda: choose_input_file(input_var),
    ).grid(row=0, column=2, padx=(8, 0))

    ttk.Label(form, text="保存目录").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
    ttk.Entry(form, textvariable=folder_var).grid(row=1, column=1, sticky="ew", pady=(10, 0))
    ttk.Button(
        form,
        text="选择目录",
        command=lambda: choose_output_folder(folder_var),
    ).grid(row=1, column=2, padx=(8, 0), pady=(10, 0))

    ttk.Label(form, text="输出文件名").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
    ttk.Entry(form, textvariable=filename_var).grid(row=2, column=1, sticky="ew", pady=(10, 0))

    form.columnconfigure(1, weight=1)

    ttk.Button(
        parent,
        text="导出 Markdown",
        command=lambda: export_to_markdown(input_var, folder_var, filename_var),
    ).pack(anchor=tk.E, pady=(14, 0))


def warn_if_pandoc_missing() -> None:
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        show_warning(
            "Pandoc 提示",
            "Pandoc 当前不可用。\n涉及 Word 转换时请先安装 Pandoc。",
        )


def build_ui(initial_tab: int = TAB_MARKDOWN_TO_WORD) -> tk.Tk:
    root = tk.Tk()
    root.title("Md2Word Toolbox")
    root.geometry("920x640")
    root.minsize(760, 540)

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True)

    tab_markdown_to_word = ttk.Frame(notebook, padding=12)
    tab_file_to_markdown = ttk.Frame(notebook, padding=12)

    notebook.add(tab_markdown_to_word, text="Markdown -> Word")
    notebook.add(tab_file_to_markdown, text="Word/PDF -> Markdown")

    build_markdown_to_word_tab(tab_markdown_to_word)
    build_file_to_markdown_tab(tab_file_to_markdown)

    if initial_tab in (TAB_MARKDOWN_TO_WORD, TAB_FILE_TO_MARKDOWN):
        notebook.select(initial_tab)

    return root


def run(initial_tab: int = TAB_MARKDOWN_TO_WORD) -> None:
    app = build_ui(initial_tab=initial_tab)
    warn_if_pandoc_missing()
    app.mainloop()


def main() -> None:
    run(initial_tab=TAB_MARKDOWN_TO_WORD)


if __name__ == "__main__":
    main()

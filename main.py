from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pypandoc


def normalize_filename(filename: str) -> str:
    name = filename.strip()
    if not name:
        raise ValueError("File name cannot be empty.")
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def convert_markdown_to_docx(markdown_text: str, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    pypandoc.convert_text(
        source=markdown_text,
        to="docx",
        format="md",
        outputfile=str(output_file),
    )


def choose_output_folder(folder_var: tk.StringVar) -> None:
    selected = filedialog.askdirectory(title="Select output folder")
    if selected:
        folder_var.set(selected)


def export_to_word(
    markdown_input: tk.Text,
    folder_var: tk.StringVar,
    filename_var: tk.StringVar,
) -> None:
    markdown_text = markdown_input.get("1.0", tk.END).strip()
    if not markdown_text:
        messagebox.showwarning("Missing content", "Please paste Markdown content first.")
        return

    folder_text = folder_var.get().strip()
    if not folder_text:
        messagebox.showwarning("Missing folder", "Please select an output folder.")
        return

    try:
        output_name = normalize_filename(filename_var.get())
    except ValueError as exc:
        messagebox.showwarning("Invalid file name", str(exc))
        return

    output_path = Path(folder_text) / output_name

    try:
        convert_markdown_to_docx(markdown_text, output_path)
    except OSError as exc:
        messagebox.showerror(
            "Pandoc not found",
            "Pandoc executable was not found. Install Pandoc first.\n\n"
            "https://pandoc.org/installing.html\n\n"
            f"Details: {exc}",
        )
        return
    except RuntimeError as exc:
        messagebox.showerror("Conversion failed", f"Failed to export Word file.\n\n{exc}")
        return

    messagebox.showinfo("Done", f"Word file exported:\n{output_path}")


def warn_if_pandoc_missing() -> None:
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        messagebox.showwarning(
            "Pandoc check",
            "Pandoc is not available right now.\n"
            "You can still edit content, but exporting needs Pandoc installed.",
        )


def build_ui() -> tk.Tk:
    root = tk.Tk()
    root.title("Markdown to Word")
    root.geometry("900x620")
    root.minsize(720, 500)

    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    title = ttk.Label(main, text="Paste Markdown and export as Word (.docx)")
    title.pack(anchor=tk.W, pady=(0, 8))

    editor_frame = ttk.Frame(main)
    editor_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    markdown_input = tk.Text(editor_frame, wrap=tk.WORD, undo=True)
    scroll = ttk.Scrollbar(editor_frame, orient=tk.VERTICAL, command=markdown_input.yview)
    markdown_input.configure(yscrollcommand=scroll.set)
    markdown_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    folder_var = tk.StringVar(value=str(Path.cwd()))
    filename_var = tk.StringVar(value="output.docx")

    output_frame = ttk.Frame(main)
    output_frame.pack(fill=tk.X, pady=(0, 8))

    ttk.Label(output_frame, text="Output folder").grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Entry(output_frame, textvariable=folder_var).grid(row=0, column=1, sticky="ew")
    ttk.Button(
        output_frame,
        text="Browse",
        command=lambda: choose_output_folder(folder_var),
    ).grid(row=0, column=2, padx=(8, 0))

    ttk.Label(output_frame, text="File name").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
    ttk.Entry(output_frame, textvariable=filename_var).grid(row=1, column=1, sticky="ew", pady=(8, 0))

    output_frame.columnconfigure(1, weight=1)

    ttk.Button(
        main,
        text="Export to Word",
        command=lambda: export_to_word(markdown_input, folder_var, filename_var),
    ).pack(anchor=tk.E)

    return root


def main() -> None:
    app = build_ui()
    warn_if_pandoc_missing()
    app.mainloop()


if __name__ == "__main__":
    main()

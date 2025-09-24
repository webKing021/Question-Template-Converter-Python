import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk


LEVEL_NUMERIC_MAP: Dict[str, str] = {
    "1": "easy",
    "2": "medium",
    "3": "hard",
    "4": "very_hard",
}

LEVEL_TEXT_MAP: Dict[str, str] = {
    "easy": "easy",
    "simple": "easy",
    "basic": "easy",
    "medium": "medium",
    "moderate": "medium",
    "modreate": "medium",  # common typo observed in sample data
    "intermediate": "medium",
    "hard": "hard",
    "difficult": "hard",
    "challenging": "hard",
    "veryhard": "very_hard",
    "extreme": "very_hard",
}

QUESTION_TYPE_MAP: Dict[str, str] = {
    "R": "multiple_choice",
    "C": "multiple_choice",
    "L": "short_answer",
}

OUTPUT_FIELDS: Sequence[str] = (
    "question_text",
    "question_type",
    "option_a",
    "option_b",
    "option_c",
    "option_d",
    "correct_answer",
    "marks",
    "difficulty_level",
    "explanation",
)

@dataclass(frozen=True)
class Theme:
    name: str
    background: str
    card: str
    surface: str
    surface_alt: str
    text_primary: str
    text_secondary: str
    primary: str
    primary_active: str
    on_primary: str
    accent: str
    accent_active: str
    on_accent: str
    tree_heading: str


THEMES: Dict[str, Theme] = {
    "Aurora Night": Theme(
        name="Aurora Night",
        background="#0b1120",
        card="#111827",
        surface="#0f172a",
        surface_alt="#1e293b",
        text_primary="#e2e8f0",
        text_secondary="#94a3b8",
        primary="#2563eb",
        primary_active="#1f85d7",
        on_primary="#f8fafc",
        accent="#38bdf8",
        accent_active="#7dd3fc",
        on_accent="#0f172a",
        tree_heading="#1e293b",
    ),
    "Solar Dawn": Theme(
        name="Solar Dawn",
        background="#f8fafc",
        card="#ffffff",
        surface="#eff6ff",
        surface_alt="#e2e8f0",
        text_primary="#0f172a",
        text_secondary="#475569",
        primary="#ef6c00",
        primary_active="#fb8c00",
        on_primary="#ffffff",
        accent="#0284c7",
        accent_active="#0ea5e9",
        on_accent="#f8fafc",
        tree_heading="#cbd5f5",
    ),
    "Emerald Forest": Theme(
        name="Emerald Forest",
        background="#06281e",
        card="#0d3b2b",
        surface="#0c3b2a",
        surface_alt="#12513b",
        text_primary="#ecfdf5",
        text_secondary="#a7f3d0",
        primary="#34d399",
        primary_active="#10b981",
        on_primary="#023020",
        accent="#2dd4bf",
        accent_active="#5eead4",
        on_accent="#023020",
        tree_heading="#146a46",
    ),
}

DEFAULT_THEME_NAME = "Aurora Night"
DEFAULT_FONT_FAMILY = "Segoe UI"
FONT_CHOICES: Sequence[str] = ("Segoe UI", "Calibri", "Verdana", "Helvetica")
FONT_SIZE_OFFSET_RANGE = (-2, 4)


@dataclass
class QuestionRecord:
    metadata_row: Dict[str, str]
    answers: List[Dict[str, str]]


def _normalize_key(key: str) -> str:
    return "".join(ch.lower() for ch in key if ch.isalnum())


def _get_field(row: Dict[str, str], *aliases: str) -> str:
    for alias in aliases:
        alias_key = _normalize_key(alias)
        for existing_key, value in row.items():
            if _normalize_key(existing_key) == alias_key:
                return (value or "").strip()
    return ""


def _determine_difficulty(level_raw: str) -> str:
    level_raw = (level_raw or "").strip()
    if not level_raw:
        return "easy"

    digits = "".join(ch for ch in level_raw if ch.isdigit())
    if digits:
        return LEVEL_NUMERIC_MAP.get(digits, "easy")

    normalized = _normalize_key(level_raw)
    return LEVEL_TEXT_MAP.get(normalized, "easy")


def _determine_question_type(code: str) -> str:
    return QUESTION_TYPE_MAP.get((code or "").strip().upper(), "multiple_choice")


def _build_output_record(question: Dict[str, str], answers: List[Dict[str, str]], warnings: List[str]) -> Dict[str, str]:
    question_text = _get_field(question, "Description", "Question")
    marks = _get_field(question, "Marks") or "1"
    level_raw = _get_field(question, "LEVEL", "Difficulty", "EASY")
    difficulty = _determine_difficulty(level_raw)

    question_type_code = _get_field(
        question,
        "QuestionTNpe",
        "QuestionType",
        "QuestionTNpe(R=Radio,C=Checkbox,L=Onelinner)",
    )
    question_type = _determine_question_type(question_type_code)

    option_fields = ["", "", "", ""]
    correct_letters: List[str] = []
    letter_sequence = ["a", "b", "c", "d"]

    if len(answers) > 4:
        warnings.append(
            f"Question '{question_text[:30]}...' has more than 4 answer options; only the first four were retained."
        )

    for idx, answer in enumerate(answers[:4]):
        option_fields[idx] = _get_field(answer, "Description", "Answer")
        is_correct = _get_field(answer, "IsRightAnswer").strip().upper() == "Y"
        if is_correct:
            correct_letters.append(letter_sequence[idx])

    correct_answer = ",".join(correct_letters)

    if question_type == "short_answer" and not correct_answer and answers:
        # Assume the first provided answer is the expected one.
        correct_answer = option_fields[0]

    return {
        "question_text": question_text,
        "question_type": question_type,
        "option_a": option_fields[0],
        "option_b": option_fields[1],
        "option_c": option_fields[2],
        "option_d": option_fields[3],
        "correct_answer": correct_answer,
        "marks": marks,
        "difficulty_level": difficulty,
        "explanation": "",
    }


def _read_question_records(input_path: Path) -> Tuple[List[QuestionRecord], List[str]]:
    warnings: List[str] = []
    questions: List[QuestionRecord] = []
    current_question: Dict[str, str] | None = None
    current_answers: List[Dict[str, str]] = []

    # Try multiple encodings to handle different file formats
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with input_path.open("r", encoding=encoding, newline="") as csv_file:
                reader = csv.DictReader(csv_file)
                for row in reader:
                    tnpe = _get_field(row, "TNpe", "Type").upper()
                    if not tnpe:
                        continue
                    if tnpe == "Q":
                        if current_question is not None:
                            questions.append(QuestionRecord(current_question, current_answers))
                        current_question = row
                        current_answers = []
                    elif tnpe == "A":
                        if current_question is None:
                            warnings.append("Encountered answer row before any question row; skipping answer.")
                            continue
                        current_answers.append(row)
                    else:
                        warnings.append(f"Unrecognized TNpe value '{tnpe}' encountered; row skipped.")
            break  # Successfully read with this encoding
        except UnicodeDecodeError:
            continue
    else:
        # If all encodings fail, raise an error
        raise UnicodeDecodeError("Failed to read CSV file with any supported encoding")

    if current_question is not None:
        questions.append(QuestionRecord(current_question, current_answers))

    return questions, warnings


def convert_question_bank(input_path: Path) -> Tuple[List[Dict[str, str]], List[str]]:
    questions, warnings = _read_question_records(input_path)

    output_rows: List[Dict[str, str]] = []
    for question_record in questions:
        output_rows.append(
            _build_output_record(question_record.metadata_row, question_record.answers, warnings)
        )

    return output_rows, warnings


def load_conversion_preview(
    input_path: Path,
) -> Tuple[List[Tuple[QuestionRecord, Dict[str, str]]], List[str]]:
    questions, warnings = _read_question_records(input_path)
    preview: List[Tuple[QuestionRecord, Dict[str, str]]] = []

    for question_record in questions:
        converted = _build_output_record(
            question_record.metadata_row, question_record.answers, warnings
        )
        preview.append((question_record, converted))

    return preview, warnings


def write_output_csv(output_path: Path, rows: Iterable[Dict[str, str]]) -> None:
    with output_path.open("w", encoding="utf-8", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=OUTPUT_FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


class ConverterGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Question Template Converter (BY KRUTARTH RAYCHURA)")
        self.root.geometry("960x720")

        self.input_path: Path | None = None
        self.preview_pairs: List[Tuple[QuestionRecord, Dict[str, str]]] = []

        self.current_theme: Theme = THEMES[DEFAULT_THEME_NAME]
        self.font_family = DEFAULT_FONT_FAMILY
        self.font_size_offset = 0

        self.base_font = tkfont.nametofont("TkDefaultFont")
        self.base_font.configure(family=self.font_family, size=self.base_font.cget("size"))
        self.base_font_size = self.base_font.cget("size")
        self.mono_base_size = tkfont.Font(font=("Consolas", 10)).cget("size")

        self._configure_styles()
        self._build_widgets()
        self._build_menu()
        self._apply_theme()
        self._update_fonts()

    def _build_widgets(self) -> None:
        self.header_frame = tk.Frame(self.root, padx=24, pady=18)
        self.header_frame.pack(fill=tk.X)

        self.header_label = tk.Label(
            self.header_frame,
            text="Question Template Converter (BY KRUTARTH RAYCHURA)",
            anchor="w",
        )
        self.header_label.pack(fill=tk.X)

        self.subtitle_label = tk.Label(
            self.header_frame,
            text="Transform your incomplete question bank into a polished quiz template with a single click.",
            anchor="w",
        )
        self.subtitle_label.pack(fill=tk.X, pady=(6, 0))

        self.main_frame = tk.Frame(self.root, padx=24, pady=12)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.control_card = ttk.Frame(self.main_frame, style="Card.TFrame", padding=(20, 16))
        self.control_card.pack(fill=tk.X)

        self.instruction_label = tk.Label(
            self.control_card,
            text=(
                "Upload the incomplete CSV to preview how each question maps into the quiz template. "
                "When you're satisfied, export the converted file ready for import."
            ),
            justify=tk.LEFT,
            wraplength=820,
        )
        self.instruction_label.pack(anchor=tk.W)

        self.button_frame = tk.Frame(self.control_card)
        self.button_frame.pack(fill=tk.X, pady=(18, 0))

        self.select_button = ttk.Button(
            self.button_frame,
            text="Select Incomplete CSV",
            command=self._choose_input_file,
            style="Primary.TButton",
        )
        self.select_button.pack(side=tk.LEFT)

        self.input_label = tk.Label(
            self.button_frame,
            text="No file selected",
            anchor="w",
        )
        self.input_label.pack(side=tk.LEFT, padx=(14, 0))

        self.convert_button = ttk.Button(
            self.control_card,
            text="Convert & Save",
            style="Accent.TButton",
            command=self._convert_and_save,
        )
        self.convert_button.pack(anchor=tk.W, pady=(16, 0))

        self.notebook = ttk.Notebook(self.main_frame, style="Modern.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(18, 12))

        self.preview_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=(16, 16))
        self.guide_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=(16, 16))
        self.notebook.add(self.preview_tab, text="Live Preview")
        self.notebook.add(self.guide_tab, text="Conversion Guide")

        self.tree_container = tk.Frame(self.preview_tab)
        self.tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("question", "type", "correct", "marks", "difficulty")
        self.preview_tree = ttk.Treeview(
            self.tree_container,
            columns=columns,
            show="headings",
            height=8,
            style="Preview.Treeview",
        )
        headings = {
            "question": "Question",
            "type": "Type",
            "correct": "Correct",
            "marks": "Marks",
            "difficulty": "Difficulty",
        }
        for key, title in headings.items():
            self.preview_tree.heading(key, text=title)
            self.preview_tree.column(key, width=160 if key == "question" else 120, anchor=tk.W)

        tree_scroll = ttk.Scrollbar(
            self.tree_container,
            orient=tk.VERTICAL,
            command=self.preview_tree.yview,
            style="Vertical.TScrollbar",
        )
        self.preview_tree.configure(yscrollcommand=tree_scroll.set)

        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 0))

        self.detail_frame = tk.LabelFrame(self.preview_tab, text="Conversion Details", padx=16, pady=12)
        self.detail_frame.pack(fill=tk.BOTH, expand=True, pady=(16, 0))

        self.detail_text = tk.Text(
            self.detail_frame,
            height=9,
            wrap=tk.WORD,
            relief=tk.FLAT,
            padx=12,
            pady=10,
        )
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        self.detail_text.configure(state=tk.DISABLED)

        self.preview_tree.bind("<<TreeviewSelect>>", lambda _: self._show_preview_details())

        self.guide_label = tk.Label(
            self.guide_tab,
            text="Input to Output Mapping",
            anchor="w",
        )
        self.guide_label.pack(fill=tk.X)

        self.guide_tree = ttk.Treeview(
            self.guide_tab,
            columns=("input", "output", "notes"),
            show="headings",
            height=8,
            style="Guide.Treeview",
        )
        self.guide_tree.heading("input", text="Incomplete Template")
        self.guide_tree.heading("output", text="Quiz Template")
        self.guide_tree.heading("notes", text="Notes")
        self.guide_tree.column("input", width=220, anchor=tk.W)
        self.guide_tree.column("output", width=160, anchor=tk.W)
        self.guide_tree.column("notes", width=340, anchor=tk.W)
        self.guide_tree.pack(fill=tk.BOTH, expand=True, pady=(12, 16))

        self._populate_conversion_guide()

        self.guide_note_label = tk.Label(
            self.guide_tab,
            text=(
                "Tip: Radio and Checkbox questions are exported as 'multiple_choice'. "
                "Checkbox questions simply accept multiple correct letters (e.g., 'a,c')."
            ),
            justify=tk.LEFT,
            wraplength=780,
        )
        self.guide_note_label.pack(fill=tk.X)

        self.log_card = ttk.Frame(self.main_frame, style="Card.TFrame", padding=(20, 16))
        self.log_card.pack(fill=tk.BOTH, expand=True)

        self.status_label = tk.Label(
            self.log_card,
            text="Activity Log",
            anchor="w",
        )
        self.status_label.pack(fill=tk.X)

        self.status_text = tk.Text(
            self.log_card,
            height=6,
            state=tk.DISABLED,
            wrap=tk.WORD,
            relief=tk.FLAT,
            padx=10,
            pady=10,
        )
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        self.footer_label = tk.Label(
            self.root,
            text="This Software is Designed & Developed by Krutarth Raychura.",
            pady=14,
        )
        self.footer_label.pack(side=tk.BOTTOM, fill=tk.X)

        self._update_detail_panel("Select a question to see how it transforms during conversion.")

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open Incomplete CSV...", command=self._choose_input_file)
        file_menu.add_command(label="Convert & Save", command=self._convert_and_save)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        menu_bar.add_cascade(label="File", menu=file_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Appearance Settings...", command=self._open_settings_dialog)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About Krutarth Raychura", command=self._show_about_dialog)
        menu_bar.add_cascade(label="Help", menu=help_menu)

    def _configure_styles(self) -> None:
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure("Card.TFrame")
        self.style.configure("Primary.TButton", font=(self.font_family, 11), padding=10)
        self.style.configure(
            "Accent.TButton",
            font=(self.font_family, 11),
            padding=(18, 10),
        )
        self.style.configure(
            "Modern.TNotebook",
            borderwidth=0,
            padding=6,
        )
        self.style.configure(
            "TNotebook.Tab",
            font=(self.font_family, 11, "bold"),
            padding=(12, 8),
        )
        self.style.configure("Preview.Treeview", rowheight=28, borderwidth=0)
        self.style.configure("Preview.Treeview.Heading", font=(self.font_family, 10, "bold"))
        self.style.configure("Guide.Treeview", rowheight=26, borderwidth=0)
        self.style.configure("Guide.Treeview.Heading", font=(self.font_family, 10, "bold"))
        self.style.configure("Vertical.TScrollbar", gripcount=0)

    def _apply_theme(self) -> None:
        theme = self.current_theme

        self.root.configure(bg=theme.background)

        # Apply background/foreground to Frame-based widgets
        self.header_frame.configure(bg=theme.background)
        self.main_frame.configure(bg=theme.background)
        self.button_frame.configure(bg=theme.card)
        self.tree_container.configure(bg=theme.card)

        for frame in (self.control_card, self.preview_tab, self.guide_tab, self.log_card):
            frame.configure(style="Card.TFrame")

        common_label_kwargs = {"bg": theme.background, "fg": theme.text_primary}
        self.header_label.configure(**common_label_kwargs, font=(self.font_family, 24, "bold"))
        self.subtitle_label.configure(
            bg=theme.background,
            fg=theme.text_secondary,
            font=(self.font_family, 12),
        )

        self.instruction_label.configure(
            bg=theme.card,
            fg=theme.text_secondary,
            font=(self.font_family, 11),
        )
        self.input_label.configure(bg=theme.card, fg=theme.text_primary, font=(self.font_family, 10, "italic"))
        self.guide_label.configure(bg=theme.card, fg=theme.text_primary, font=(self.font_family, 14, "bold"))
        self.guide_note_label.configure(bg=theme.card, fg=theme.text_secondary, font=(self.font_family, 10))
        self.status_label.configure(bg=theme.card, fg=theme.text_primary, font=(self.font_family, 12, "bold"))
        self.footer_label.configure(bg=theme.background, fg=theme.text_secondary)

        self.detail_frame.configure(bg=theme.card, fg=theme.text_primary)

        self.detail_text.configure(bg=theme.surface_alt, fg=theme.text_primary, font=("Consolas", self.mono_base_size + self.font_size_offset))
        self.status_text.configure(bg=theme.surface_alt, fg=theme.text_primary, font=("Consolas", self.mono_base_size + self.font_size_offset))

        self.style.configure("Card.TFrame", background=theme.card)
        self.style.configure(
            "Primary.TButton",
            background=theme.primary,
            foreground=theme.on_primary,
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", theme.primary_active)],
            foreground=[("disabled", theme.text_secondary)],
        )
        self.style.configure(
            "Accent.TButton",
            background=theme.accent,
            foreground=theme.on_accent,
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", theme.accent_active)],
            foreground=[("disabled", theme.text_secondary)],
        )
        self.style.configure("Modern.TNotebook", background=theme.background)
        self.style.configure(
            "TNotebook.Tab",
            background=theme.card,
            foreground=theme.text_secondary,
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", theme.surface_alt)],
            foreground=[("selected", theme.text_primary)],
        )
        self.style.configure(
            "Preview.Treeview",
            background=theme.surface,
            fieldbackground=theme.surface,
            foreground=theme.text_primary,
        )
        self.style.configure(
            "Preview.Treeview.Heading",
            background=theme.tree_heading,
            foreground=theme.text_primary,
        )
        self.style.configure(
            "Guide.Treeview",
            background=theme.surface,
            fieldbackground=theme.surface,
            foreground=theme.text_primary,
        )
        self.style.configure(
            "Guide.Treeview.Heading",
            background=theme.tree_heading,
            foreground=theme.text_primary,
        )

    def _show_about_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("About Converter")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=self.current_theme.background)
        
        # Center the dialog
        dialog.geometry("400x500")
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"400x500+{x}+{y}")
        
        main_frame = tk.Frame(dialog, bg=self.current_theme.background, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Load and display the logo image
        try:
            logo_path = Path(__file__).parent / "logo.jpg"
            if logo_path.exists():
                pil_image = Image.open(logo_path)
                # Resize image to fit nicely in the dialog
                pil_image = pil_image.resize((150, 150), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(pil_image)
                
                logo_label = tk.Label(main_frame, image=photo, bg=self.current_theme.background)
                logo_label.image = photo  # Keep a reference to prevent garbage collection
                logo_label.pack(pady=(0, 20))
        except Exception:
            # If image loading fails, just skip it
            pass
        
        # App title
        title_label = tk.Label(
            main_frame,
            text="KRUTARTH RAYCHURA",
            font=(self.font_family, 16, "bold"),
            bg=self.current_theme.background,
            fg=self.current_theme.text_primary,
        )
        title_label.pack(pady=(0, 5))
        
        # Version
        version_label = tk.Label(
            main_frame,
            text="Version 1.0",
            font=(self.font_family, 11),
            bg=self.current_theme.background,
            fg=self.current_theme.text_secondary,
        )
        version_label.pack(pady=(0, 15))
        
        # Description
        description_text = (
            "Convert the incomplete question bank into the quiz import template "
            "with live previews and customizable appearance."
        )
        description_label = tk.Label(
            main_frame,
            text=description_text,
            font=(self.font_family, 10),
            bg=self.current_theme.background,
            fg=self.current_theme.text_primary,
            wraplength=340,
            justify=tk.CENTER,
        )
        description_label.pack(pady=(0, 20))
        
        # Developer credit
        credit_label = tk.Label(
            main_frame,
            text="This Software is Designed & Developed by\nKrutarth Raychura",
            font=(self.font_family, 11, "bold"),
            bg=self.current_theme.background,
            fg=self.current_theme.accent,
            justify=tk.CENTER,
        )
        credit_label.pack(pady=(0, 20))
        
        # Close button
        close_button = ttk.Button(
            main_frame,
            text="Close",
            command=dialog.destroy,
            style="Primary.TButton",
        )
        close_button.pack()

    def _open_settings_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Appearance Settings")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=self.current_theme.background)

        ttk.Label(dialog, text="Theme:").grid(row=0, column=0, padx=12, pady=(12, 8), sticky="w")
        theme_var = tk.StringVar(value=self.current_theme.name)
        theme_combo = ttk.Combobox(dialog, textvariable=theme_var, values=list(THEMES.keys()), state="readonly")
        theme_combo.grid(row=0, column=1, padx=12, pady=(12, 8), sticky="ew")

        ttk.Label(dialog, text="Font Family:").grid(row=1, column=0, padx=12, pady=8, sticky="w")
        font_var = tk.StringVar(value=self.font_family)
        font_combo = ttk.Combobox(dialog, textvariable=font_var, values=list(FONT_CHOICES), state="readonly")
        font_combo.grid(row=1, column=1, padx=12, pady=8, sticky="ew")

        ttk.Label(dialog, text="Font Size Offset:").grid(row=2, column=0, padx=12, pady=8, sticky="w")
        size_var = tk.IntVar(value=self.font_size_offset)
        size_spin = ttk.Spinbox(
            dialog,
            from_=FONT_SIZE_OFFSET_RANGE[0],
            to=FONT_SIZE_OFFSET_RANGE[1],
            textvariable=size_var,
            width=6,
        )
        size_spin.grid(row=2, column=1, padx=12, pady=8, sticky="w")

        dialog.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(16, 12))

        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(
            button_frame,
            text="Apply",
            command=lambda: self._apply_settings_changes(dialog, theme_var.get(), font_var.get(), size_var.get()),
        ).pack(side=tk.RIGHT)

    def _populate_conversion_guide(self) -> None:
        guide_rows = [
            (
                "TNpe = Q → Description",
                "question_text",
                "Question rows keep their description as the main text.",
            ),
            (
                "QuestionTNpe (R/C/L)",
                "question_type",
                "R & C become 'multiple_choice'; L becomes 'short_answer'.",
            ),
            (
                "TNpe = A → Description", "option_a-d", "Answer descriptions populate up to four options in order.",
            ),
            (
                "IsRightAnswer = Y", "correct_answer", "Correct answers are stored as letters a–d (multiple letters for checkbox).",
            ),
            (
                "Marks", "marks", "Marks transfer directly to the output.",
            ),
            (
                "LEVEL", "difficulty_level", "Numeric levels map to easy/medium/hard/very_hard.",
            ),
            (
                "IsImage / ImagePath", "ignored", "Image information is not required and left out.",
            ),
        ]

        for item in self.guide_tree.get_children():
            self.guide_tree.delete(item)

        for row in guide_rows:
            self.guide_tree.insert("", tk.END, values=row)

    def _choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Incomplete Question Bank CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if file_path:
            self.input_path = Path(file_path)
            self._log(f"Selected input file: {self.input_path}")
            self.input_label.config(text=self.input_path.name)
            self._load_preview()

    def _convert_and_save(self) -> None:
        if self.input_path is None:
            messagebox.showwarning("No input file", "Please select an input CSV file first.")
            return

        try:
            rows, warnings = convert_question_bank(self.input_path)
            if not rows:
                messagebox.showwarning(
                    "No questions found",
                    "The selected file did not contain any questions to convert.",
                )
                return

            default_name = f"{self.input_path.stem}_converted.csv"
            output_file = filedialog.asksaveasfilename(
                title="Save Converted CSV",
                defaultextension=".csv",
                initialfile=default_name,
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
            if not output_file:
                return

            output_path = Path(output_file)
            write_output_csv(output_path, rows)

            messagebox.showinfo(
                "Conversion complete",
                f"Successfully saved converted CSV to:\n{output_path}",
            )

            self._log("Conversion successful.")
            for warning in warnings:
                self._log(f"Warning: {warning}")

            self._load_preview(show_alerts=False)
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Conversion failed", str(exc))
            self._log(f"Error: {exc}")

    def _load_preview(self, show_alerts: bool = True) -> None:
        if self.input_path is None:
            return

        try:
            preview_pairs, warnings = load_conversion_preview(self.input_path)
        except Exception as exc:  # pylint: disable=broad-except
            if show_alerts:
                messagebox.showerror("Preview failed", str(exc))
            self._log(f"Error generating preview: {exc}")
            return

        self.preview_pairs = preview_pairs
        self._refresh_preview_tree()

        self._log(
            f"Loaded {len(self.preview_pairs)} question(s) from '{self.input_path.name}' for preview."
        )
        for warning in warnings:
            self._log(f"Warning: {warning}")

        if self.preview_pairs:
            first_item = self.preview_tree.get_children()
            if first_item:
                self.preview_tree.selection_set(first_item[0])
                self.preview_tree.focus(first_item[0])
                self._show_preview_details()
        else:
            self._update_detail_panel("No questions were detected in the selected file.")

    def _refresh_preview_tree(self) -> None:
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        for index, (_, converted) in enumerate(self.preview_pairs):
            correct = converted["correct_answer"].replace(",", ", ") or "—"
            self.preview_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    converted["question_text"],
                    converted["question_type"],
                    correct,
                    converted["marks"],
                    converted["difficulty_level"],
                ),
            )

    def _show_preview_details(self) -> None:
        selection = self.preview_tree.selection()
        if not selection:
            return

        index = int(selection[0])
        if index >= len(self.preview_pairs):
            return

        question_record, converted = self.preview_pairs[index]

        question_type_code = _get_field(
            question_record.metadata_row,
            "QuestionTNpe",
            "QuestionType",
            "QuestionTNpe(R=Radio,C=Checkbox,L=Onelinner)",
        )
        question_type_code = question_type_code.upper() or "R"

        answers_lines = []
        for idx, answer in enumerate(question_record.answers, start=1):
            flag = _get_field(answer, "IsRightAnswer").upper() == "Y"
            marker = "[Correct]" if flag else "[ ]"
            description = _get_field(answer, "Description")
            answers_lines.append(f"  {idx}. {marker} {description}")

        converted_lines = [
            f"Question Text  : {converted['question_text']}",
            f"Question Type  : {converted['question_type']} (source code: {question_type_code or 'N/A'})",
            f"Marks          : {converted['marks']}",
            f"Difficulty     : {converted['difficulty_level']}",
            f"Correct Answer : {converted['correct_answer'] or '—'}",
            "Options        :",
            f"  A. {converted['option_a']}",
            f"  B. {converted['option_b']}",
            f"  C. {converted['option_c']}",
            f"  D. {converted['option_d']}",
        ]

        detail_text = "Input Question Row:\n" + _get_field(
            question_record.metadata_row,
            "Description",
            "Question",
        )
        detail_text += "\n\nAnswer Options:\n" + ("\n".join(answers_lines) or "  (none)")
        detail_text += "\n\nConverted Output:\n" + "\n".join(converted_lines)

        self._update_detail_panel(detail_text)

    def _update_detail_panel(self, message: str) -> None:
        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, message)
        self.detail_text.configure(state=tk.DISABLED)

    def _apply_settings_changes(self, dialog: tk.Toplevel, theme_name: str, font_family: str, size_offset: int) -> None:
        self.current_theme = THEMES.get(theme_name, self.current_theme)
        if font_family in FONT_CHOICES:
            self.font_family = font_family

        size_offset = max(FONT_SIZE_OFFSET_RANGE[0], min(FONT_SIZE_OFFSET_RANGE[1], size_offset))
        self.font_size_offset = size_offset

        self._update_fonts()
        self._apply_theme()
        dialog.destroy()

    def _update_fonts(self) -> None:
        base_size = tkfont.Font().cget("size") + self.font_size_offset
        for widget in self.root.winfo_children():
            self._update_widget_font(widget, base_size)

    def _update_widget_font(self, widget: tk.Widget, base_size: int) -> None:
        if isinstance(widget, tk.Label):
            current_font = tkfont.Font(font=widget.cget("font"))
            weight = "bold" if current_font.actual("weight") == "bold" else "normal"
            slant = "italic" if current_font.actual("slant") == "italic" else "roman"
            widget.configure(font=(self.font_family, base_size, weight, slant))
        elif isinstance(widget, tk.Button):
            widget.configure(font=(self.font_family, base_size))
        elif isinstance(widget, tk.Text):
            widget.configure(font=("Consolas", max(8, base_size)))

        for child in widget.winfo_children():
            self._update_widget_font(child, base_size)

    def _log(self, message: str) -> None:
        self.status_text.configure(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.configure(state=tk.DISABLED)
        self.status_text.see(tk.END)


def launch_gui() -> None:
    root = tk.Tk()
    ConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()

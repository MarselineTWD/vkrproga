# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import tkinter as tk
import html
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
import psycopg
from dateutil.relativedelta import relativedelta
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

try:
    from .analysis import summarize_ros, t_test_one_sample
    from .models import Enterprise, FinancialRecord, FinancialReport
    from .repository import PostgresRepository
except ImportError:
    # Support direct script run: `python rentability/ui.py`
    from analysis import summarize_ros, t_test_one_sample
    from models import Enterprise, FinancialRecord, FinancialReport
    from repository import PostgresRepository

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


TABLE_COLUMNS = (
    "Дата",
    "Выручка, ₽",
    "Себестоимость, ₽",
    "Пост. издержки, ₽",
    "Перем. издержки, ₽",
    "Налог, ₽",
    "Чистая прибыль, ₽",
    "ROS, %",
)

IMPORT_COLUMN_ALIASES = {
    "enterprise": "enterprise_name",
    "enterprise_name": "enterprise_name",
    "enterprise_title": "enterprise_name",
    "name": "enterprise_name",
    "предприятие": "enterprise_name",
    "название": "enterprise_name",
    "date": "period_date",
    "period_date": "period_date",
    "period": "period_date",
    "report_date": "period_date",
    "дата": "period_date",
    "revenue": "revenue",
    "income": "revenue",
    "sales": "revenue",
    "выручка": "revenue",
    "cost": "cost",
    "costs": "cost",
    "cost_price": "cost",
    "prime_cost": "cost",
    "себестоимость": "cost",
    "fixed_expenses": "fixed_expenses",
    "fixed_costs": "fixed_expenses",
    "fixed": "fixed_expenses",
    "постоянные издержки": "fixed_expenses",
    "пост_издержки": "fixed_expenses",
    "variable_expenses": "variable_expenses",
    "variable_costs": "variable_expenses",
    "variable": "variable_expenses",
    "переменные издержки": "variable_expenses",
    "перем_издержки": "variable_expenses",
    "tax": "tax",
    "taxes": "tax",
    "налог": "tax",
}

RECORD_FIELDS = [
    ("Дата (ГГГГ-ММ-ДД):", "period_date"),
    ("Выручка, ₽:", "revenue"),
    ("Себестоимость, ₽:", "cost"),
    ("Постоянные издержки, ₽:", "fixed_expenses"),
    ("Переменные издержки, ₽:", "variable_expenses"),
    ("Налог, ₽:", "tax"),
]

FIELD_DISPLAY_NAMES = {
    "revenue": "Выручка",
    "cost": "Себестоимость",
    "fixed_expenses": "Постоянные издержки",
    "variable_expenses": "Переменные издержки",
    "tax": "Налог",
}


COLOR_BG = "#f3f7fc"
COLOR_SURFACE = "#ffffff"
COLOR_SURFACE_ALT = "#f6fafe"
COLOR_TEXT = "#1e3550"
COLOR_MUTED = "#5f7792"
COLOR_PRIMARY = "#2f6fb6"
COLOR_PRIMARY_HOVER = "#245f9f"
COLOR_BORDER = "#d3e1ef"
COLOR_BANNER_TEXT = "#1e3d5f"
COLOR_BANNER_SUBTEXT = "#5d7592"
COLOR_BANNER_OVERLAY = "#e8f1fb"
APP_TITLE = "ОРМП: Оценка рентабельности малого предприятия"

ASSET_DIR = Path(__file__).resolve().parent / "assets"
BACKGROUND_IMAGE_PATH = ASSET_DIR / "ui_background_soft.png"
HAS_CUSTOM_TK = ctk is not None
AppWindowBase = ctk.CTk if HAS_CUSTOM_TK else tk.Tk


class RentabilityAnalysisApp(AppWindowBase):
    def __init__(self, repository: PostgresRepository | None = None):
        super().__init__()
        if HAS_CUSTOM_TK:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        self.geometry("1400x800")
        self.minsize(1220, 760)
        if HAS_CUSTOM_TK:
            self.configure(fg_color=COLOR_BG)
        else:
            self.configure(bg=COLOR_BG)

        self.repository = repository or PostgresRepository()
        try:
            self.repository.initialize()
        except psycopg.Error as exc:
            self._show_styled_error_dialog(
                "Ошибка PostgreSQL",
                "Не удалось подключиться к PostgreSQL.\n"
                "Проверьте DATABASE_URL или параметры подключения.\n\n"
                f"Текст ошибки: {exc}",
            )
            self.destroy()
            return

        self.enterprise_by_name: dict[str, Enterprise] = {}
        self.current_enterprise: Enterprise | None = None
        self.current_records: list[FinancialRecord] = []
        self.analysis_result: dict | None = None
        self.current_report: FinancialReport | None = None
        self.selected_record_id: int | None = None
        self.modal_overlay: tk.Toplevel | None = None

        self.enterprise_var = tk.StringVar()
        self.period_mode_var = tk.StringVar(value="Все данные")
        self.period_start_var = tk.StringVar()
        self.period_end_var = tk.StringVar()
        self.target_ros_var = tk.StringVar(value="10.0")
        self.alpha_var = tk.StringVar(value="0.05")
        self.graph_tick_mode_var = tk.StringVar(value="Автоматически")
        self.quick_recommendation_var = tk.StringVar(value="Практический вывод: —")

        self.fig_metrics: Figure | None = None
        self.canvas_metrics: FigureCanvasTkAgg | None = None
        self.metrics_axes: tuple | None = None
        self.drag_state: dict[str, float] | None = None
        self.graph_widget: tk.Widget | None = None
        self.suppress_tree_event = False
        self.graph_animation_after_id: str | None = None
        self.background_canvas: tk.Canvas | None = None
        self.background_image: tk.PhotoImage | None = None
        self.background_image_id: int | None = None

        self._setup_background_visual()
        self._setup_styles()
        self.create_widgets()
        self.update_enterprise_list()
        self.after(20, self._open_maximized)
        self.after(30, self._animate_window_open)

    def _open_maximized(self) -> None:
        try:
            self.state("zoomed")
            return
        except tk.TclError:
            pass
        try:
            self.attributes("-zoomed", True)
            return
        except tk.TclError:
            pass
        width = self.winfo_screenwidth()
        height = self.winfo_screenheight()
        self.geometry(f"{width}x{height}+0+0")

    def _setup_background_visual(self) -> None:
        self.background_canvas = tk.Canvas(self, bd=0, highlightthickness=0, relief=tk.FLAT, bg=COLOR_BG)
        self.background_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)

        image_path = self._ensure_background_image()
        if image_path is not None:
            try:
                self.background_image = tk.PhotoImage(file=str(image_path))
            except tk.TclError:
                self.background_image = None

        if self.background_image is not None:
            self.background_image_id = self.background_canvas.create_image(
                0,
                0,
                image=self.background_image,
                anchor="center",
            )
        self.bind("<Configure>", self._on_background_resize, add="+")

    def _ensure_background_image(self) -> Path | None:
        if BACKGROUND_IMAGE_PATH.exists():
            return BACKGROUND_IMAGE_PATH

        try:
            ASSET_DIR.mkdir(parents=True, exist_ok=True)
            width, height = 2400, 1400
            x_grad = np.linspace(0.0, 1.0, width)
            y_grad = np.linspace(0.0, 1.0, height)
            x_mesh, y_mesh = np.meshgrid(x_grad, y_grad)

            glow_a = np.exp(-((x_mesh - 0.18) ** 2 + (y_mesh - 0.18) ** 2) / 0.06)
            glow_b = np.exp(-((x_mesh - 0.82) ** 2 + (y_mesh - 0.78) ** 2) / 0.09)
            wave = 0.5 + 0.5 * np.sin((x_mesh * 0.8 + y_mesh * 0.55) * np.pi * 1.6)

            red = 242 + 6 * x_mesh + 10 * glow_a + 6 * glow_b
            green = 247 + 5 * y_mesh + 9 * glow_a + 5 * wave
            blue = 253 + 9 * x_mesh + 11 * glow_b + 5 * wave

            image = np.clip(np.dstack((red, green, blue)), 0, 255).astype(np.uint8)
            plt.imsave(BACKGROUND_IMAGE_PATH, image)
            return BACKGROUND_IMAGE_PATH
        except Exception:
            return None

    def _on_background_resize(self, _event=None) -> None:
        if not self.background_canvas:
            return
        width = self.winfo_width()
        height = self.winfo_height()
        if width <= 1 or height <= 1:
            return
        self.background_canvas.configure(width=width, height=height)
        if self.background_image_id is not None:
            self.background_canvas.coords(self.background_image_id, width // 2, height // 2)

    def _create_hero_banner(self, parent: ttk.Frame) -> None:
        banner = tk.Canvas(
            parent,
            height=120,
            bg=COLOR_BANNER_OVERLAY,
            bd=0,
            highlightthickness=0,
            relief=tk.FLAT,
        )
        banner.pack(fill=tk.X, padx=8, pady=(8, 10))

        banner.create_rectangle(0, 0, 6000, 120, fill=COLOR_BANNER_OVERLAY, outline="")
        banner.create_text(
            26,
            40,
            text=APP_TITLE,
            anchor="w",
            fill=COLOR_BANNER_TEXT,
            font=("Segoe UI Semibold", 18),
        )
        banner.create_text(
            26,
            78,
            text="Автоматизированное рабочее место анализа финансовых показателей",
            anchor="w",
            fill=COLOR_BANNER_SUBTEXT,
            font=("Segoe UI", 11),
        )

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("CardInner.TFrame", background=COLOR_SURFACE)
        style.configure("TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure(
            "Card.TLabelframe",
            background=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            borderwidth=2,
            relief="solid",
            padding=10,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=COLOR_BG,
            foreground=COLOR_TEXT,
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "TButton",
            background=COLOR_PRIMARY,
            foreground="#ffffff",
            padding=(14, 8),
            borderwidth=0,
            focusthickness=0,
            font=("Segoe UI Semibold", 10),
        )
        style.map("TButton", background=[("active", COLOR_PRIMARY_HOVER), ("disabled", "#9cb4cd")])
        style.configure(
            "Dialog.TFrame",
            background=COLOR_SURFACE,
            relief="flat",
        )
        style.configure(
            "Dialog.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Dialog.TButton",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 7),
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "Dialog.TButton",
            background=[("active", "#e8f1fb"), ("disabled", "#e0e8f2")],
            foreground=[("disabled", "#8ca2bb")],
        )
        style.configure(
            "QuickResult.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_PRIMARY,
            font=("Segoe UI Semibold", 13),
            padding=(14, 6),
        )
        style.configure(
            "TEntry",
            fieldbackground=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=6,
        )
        style.configure(
            "TCombobox",
            fieldbackground=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            selectbackground=COLOR_SURFACE,
            selectforeground=COLOR_TEXT,
            padding=5,
        )
        style.map("TCombobox", fieldbackground=[("readonly", COLOR_SURFACE)])
        style.configure(
            "Treeview",
            background=COLOR_SURFACE,
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            borderwidth=0,
            relief="flat",
            rowheight=31,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_TEXT,
            relief="flat",
            borderwidth=0,
            font=("Segoe UI Semibold", 10),
            padding=(5, 5),
        )
        style.map("Treeview.Heading", background=[("active", "#e4effa")])
        style.configure(
            "Vertical.TScrollbar",
            gripcount=0,
            background="#d9e6f4",
            troughcolor="#f2f7fd",
            bordercolor=COLOR_BORDER,
            arrowcolor="#5f7792",
            lightcolor="#d9e6f4",
            darkcolor="#d9e6f4",
        )
        style.map("Vertical.TScrollbar", background=[("active", "#c9dbef")])
        style.configure(
            "Horizontal.TScrollbar",
            gripcount=0,
            background="#d9e6f4",
            troughcolor="#f2f7fd",
            bordercolor=COLOR_BORDER,
            arrowcolor="#5f7792",
            lightcolor="#d9e6f4",
            darkcolor="#d9e6f4",
        )
        style.map("Horizontal.TScrollbar", background=[("active", "#c9dbef")])
        style.configure("Main.TPanedwindow", background=COLOR_BG, sashwidth=10)

    def _animate_window_open(self) -> None:
        try:
            self.attributes("-alpha", 0.0)
        except tk.TclError:
            return
        self._animate_alpha(self, target_alpha=1.0, duration_ms=240)

    def _animate_alpha(
        self,
        window: tk.Toplevel | tk.Tk,
        *,
        target_alpha: float,
        duration_ms: int = 180,
        steps: int = 10,
    ) -> None:
        try:
            current_alpha = float(window.attributes("-alpha"))
        except tk.TclError:
            return
        step_ms = max(10, duration_ms // max(steps, 1))
        delta = (target_alpha - current_alpha) / max(steps, 1)

        def tick(step: int = 1) -> None:
            if not window.winfo_exists():
                return
            next_alpha = current_alpha + delta * step
            try:
                window.attributes("-alpha", max(0.0, min(target_alpha, next_alpha)))
            except tk.TclError:
                return
            if step < steps:
                self.after(step_ms, lambda: tick(step + 1))

        tick()

    def create_widgets(self) -> None:
        root_frame = ttk.Frame(self, style="App.TFrame")
        root_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        root_frame.lift()

        main_paned = ttk.PanedWindow(root_frame, orient=tk.HORIZONTAL, style="Main.TPanedwindow")
        main_paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        left_panel = ttk.Frame(main_paned, style="App.TFrame")
        right_panel = ttk.Frame(main_paned, style="App.TFrame")
        main_paned.add(left_panel, weight=5)
        main_paned.add(right_panel, weight=6)

        self._create_table_panel(left_panel)
        self._create_settings_panel(right_panel)
        self._create_graph_panel(right_panel)

    def _create_table_panel(self, parent: ttk.Frame) -> None:
        table_frame = ttk.LabelFrame(parent, text="Финансовые данные за период")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.configure(style="Card.TLabelframe")

        self.tree = ttk.Treeview(table_frame, columns=TABLE_COLUMNS, show="headings", height=12)
        for column, width in zip(TABLE_COLUMNS, [95, 130, 140, 145, 145, 100, 140, 70]):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, minwidth=width, anchor=tk.CENTER, stretch=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("even", background="#eef5fd")

        y_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        data_actions_frame = ttk.Frame(parent, style="App.TFrame")
        data_actions_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(data_actions_frame, text="Добавить данные", command=self.add_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_actions_frame, text="Редактировать", command=self.edit_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_actions_frame, text="Удалить данные", command=self.delete_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_actions_frame, text="Импорт", command=self.import_data).pack(side=tk.LEFT, padx=2)

        results_frame = ttk.LabelFrame(parent, text="Статистические показатели")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        results_frame.configure(style="Card.TLabelframe")

        results_actions = ttk.Frame(results_frame, style="CardInner.TFrame")
        results_actions.pack(fill=tk.X, padx=5, pady=(5, 0))
        ttk.Button(
            results_actions,
            text="Рассчитать показатели",
            command=self.test_hypothesis,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Label(
            results_actions,
            textvariable=self.quick_recommendation_var,
            style="QuickResult.TLabel",
        ).pack(side=tk.LEFT, padx=(10, 2))

        self.results_text = tk.Text(
            results_frame,
            height=11,
            font=("Segoe UI", 11),
            state=tk.DISABLED,
            wrap=tk.WORD,
            padx=8,
            pady=8,
        )
        self.results_text.configure(
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_TEXT,
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER,
            insertbackground=COLOR_TEXT,
        )
        results_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scrollbar.set)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_settings_panel(self, parent: ttk.Frame) -> None:
        settings_frame = ttk.LabelFrame(parent, text="Параметры анализа")
        settings_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        settings_frame.configure(style="Card.TLabelframe")

        enterprise_frame = ttk.Frame(settings_frame, style="CardInner.TFrame")
        enterprise_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(enterprise_frame, text="Предприятие:").pack(side=tk.LEFT)
        self.enterprise_combo = ttk.Combobox(
            enterprise_frame,
            textvariable=self.enterprise_var,
            state="readonly",
            width=30,
        )
        self.enterprise_combo.pack(side=tk.LEFT, padx=5)
        self.enterprise_combo.bind("<<ComboboxSelected>>", self.on_enterprise_change)
        self._setup_combobox_behavior(self.enterprise_combo)
        ttk.Button(
            enterprise_frame,
            text="Добавить предприятие",
            command=self.add_enterprise,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            enterprise_frame,
            text="Удалить предприятие",
            command=self.delete_enterprise,
        ).pack(side=tk.LEFT, padx=2)

        row1 = ttk.Frame(settings_frame, style="CardInner.TFrame")
        row1.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row1, text="Дата анализа:").pack(side=tk.LEFT)
        ttk.Label(row1, text=datetime.now().strftime("%d.%m.%Y")).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="Период:").pack(side=tk.LEFT, padx=(20, 5))
        period_mode_combo = ttk.Combobox(
            row1,
            textvariable=self.period_mode_var,
            state="readonly",
            width=12,
            values=("Все данные", "Период"),
        )
        period_mode_combo.pack(side=tk.LEFT, padx=5)
        period_mode_combo.bind("<<ComboboxSelected>>", self.on_period_mode_change)
        self._setup_combobox_behavior(period_mode_combo)

        self.period_range_frame = ttk.Frame(row1, style="CardInner.TFrame", width=290, height=32)
        self.period_range_frame.pack(side=tk.LEFT, padx=(10, 0))
        self.period_range_frame.pack_propagate(False)
        self.period_from_label = ttk.Label(self.period_range_frame, text="с")
        self.period_from_entry = ttk.Entry(self.period_range_frame, textvariable=self.period_start_var, width=12)
        self.period_to_label = ttk.Label(self.period_range_frame, text="по")
        self.period_to_entry = ttk.Entry(self.period_range_frame, textvariable=self.period_end_var, width=12)
        self._update_period_mode_ui()

        row2 = ttk.Frame(settings_frame, style="CardInner.TFrame")
        row2.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row2, text="Целевой ROS (%):").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.target_ros_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Уровень α:").pack(side=tk.LEFT, padx=(20, 5))
        ttk.Entry(row2, textvariable=self.alpha_var, width=6).pack(side=tk.LEFT, padx=5)

        button_frame = ttk.Frame(settings_frame, style="CardInner.TFrame")
        button_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        ttk.Button(button_frame, text="Вывести данные", command=self.show_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Сохранить отчёт", command=self.save_report).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Отчёты", command=self.show_saved_reports).pack(side=tk.LEFT, padx=2)

    def _create_graph_panel(self, parent: ttk.Frame) -> None:
        graph_frame = ttk.LabelFrame(parent, text="")
        graph_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        graph_frame.configure(style="Card.TLabelframe")

        graph_controls = ttk.Frame(graph_frame, style="CardInner.TFrame")
        graph_controls.pack(fill=tk.X, padx=8, pady=(6, 0))
        ttk.Label(graph_controls, text="Подписи месяцев:").pack(side=tk.LEFT)
        tick_mode_combo = ttk.Combobox(
            graph_controls,
            textvariable=self.graph_tick_mode_var,
            state="readonly",
            width=20,
            values=("Автоматически", "Каждый месяц", "Через 2 месяца"),
        )
        tick_mode_combo.pack(side=tk.LEFT, padx=6)
        tick_mode_combo.bind("<<ComboboxSelected>>", self.on_graph_tick_mode_change)
        self._setup_combobox_behavior(tick_mode_combo)


        self.fig_metrics = Figure(figsize=(8, 6), dpi=100)
        self.canvas_metrics = FigureCanvasTkAgg(self.fig_metrics, master=graph_frame)
        self.graph_widget = self.canvas_metrics.get_tk_widget()
        self.graph_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=(4, 5))
        self.graph_widget.configure(cursor="hand2", bg=COLOR_SURFACE)
        self.canvas_metrics.mpl_connect("scroll_event", self.on_graph_scroll)
        self.canvas_metrics.mpl_connect("button_press_event", self.on_graph_press)
        self.canvas_metrics.mpl_connect("motion_notify_event", self.on_graph_drag)
        self.canvas_metrics.mpl_connect("button_release_event", self.on_graph_release)
        self.canvas_metrics.mpl_connect("figure_leave_event", self.on_graph_leave)

    def update_enterprise_list(self) -> None:
        enterprises = self.repository.list_enterprises()
        self.enterprise_by_name = {enterprise.name: enterprise for enterprise in enterprises}
        names = [enterprise.name for enterprise in enterprises]
        self.enterprise_combo["values"] = names
        if names and self.enterprise_var.get() not in self.enterprise_by_name:
            self.enterprise_var.set(names[0])
        enterprise = self.enterprise_by_name.get(self.enterprise_var.get())
        if enterprise:
            self.set_period_to_full_range(enterprise)

    def parse_date_string(self, date_str: str) -> date | None:
        value = date_str.strip()
        try:
            if "." in value:
                return datetime.strptime(value, "%d.%m.%Y").date()
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None

    @staticmethod
    def _ensure_not_future_date(target_date: date, field_name: str = "Дата") -> None:
        if target_date > date.today():
            raise ValueError(f"{field_name} не может быть в будущем")

    def get_selected_enterprise(self) -> Enterprise | None:
        enterprise = self.enterprise_by_name.get(self.enterprise_var.get())
        if not enterprise:
            self._show_styled_warning_dialog("Ошибка", "Выберите предприятие")
            return None
        return enterprise

    def clear_results(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.selected_record_id = None
        self.current_records = []
        self.analysis_result = None
        self.current_report = None
        self.quick_recommendation_var.set("Практический вывод: —")

        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        self.results_text.config(state=tk.DISABLED)

        self.clear_graphs_only()

    def clear_statistics(self) -> None:
        self.analysis_result = None
        self.current_report = None
        self.quick_recommendation_var.set("Практический вывод: —")
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        self.results_text.config(state=tk.DISABLED)

    def render_statistics(self, lines: list[str]) -> None:
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        self.results_text.insert(tk.END, "\n".join(lines))
        self.results_text.config(state=tk.DISABLED)

    def set_period_to_full_range(self, enterprise: Enterprise) -> None:
        start_date, end_date = self.repository.get_record_date_bounds(enterprise.id)
        self.period_start_var.set(start_date.strftime("%Y-%m-%d") if start_date else "")
        self.period_end_var.set(end_date.strftime("%Y-%m-%d") if end_date else "")

    def ensure_period_includes(self, target_date: date) -> None:
        start_date = self.parse_date_string(self.period_start_var.get())
        end_date = self.parse_date_string(self.period_end_var.get())

        if start_date is None or target_date < start_date:
            self.period_start_var.set(target_date.strftime("%Y-%m-%d"))
        if end_date is None or target_date > end_date:
            self.period_end_var.set(target_date.strftime("%Y-%m-%d"))

    def refresh_current_view(self, *, rerun_analysis: bool = True) -> None:
        self.update_enterprise_list()
        self.show_data()
        if rerun_analysis and self.analysis_result and len(self.current_records) >= 2:
            self.test_hypothesis()

    def _create_modal_overlay(self) -> tk.Toplevel:
        self.update_idletasks()
        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.transient(self)
        overlay.configure(bg="black")
        try:
            overlay.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        overlay.geometry(
            f"{self.winfo_width()}x{self.winfo_height()}+{self.winfo_rootx()}+{self.winfo_rooty()}"
        )
        return overlay

    def _center_dialog(self, dialog: tk.Toplevel) -> None:
        self.update_idletasks()
        dialog.update_idletasks()
        x_pos = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y_pos = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x_pos, 0)}+{max(y_pos, 0)}")

    def _style_dialog(self, dialog: tk.Toplevel, *, size: str, resizable: tuple[bool, bool] = (False, False)) -> None:
        dialog.geometry(size)
        dialog.resizable(*resizable)
        dialog.configure(bg=COLOR_SURFACE, highlightthickness=1, highlightbackground=COLOR_BORDER, highlightcolor=COLOR_BORDER)

    def _show_styled_info_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        self._style_dialog(dialog, size="430x170")

        body = ttk.Frame(dialog, style="Dialog.TFrame")
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        ttk.Label(body, text=title, style="Dialog.TLabel", font=("Segoe UI Semibold", 12)).pack(anchor="w", pady=(0, 8))
        text_widget = tk.Text(
            body,
            height=3,
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            padx=2,
            pady=2,
            font=("Segoe UI", 10),
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", message)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(
            body,
            text="OK",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(anchor="e", pady=(10, 0))

        self._show_modal_dialog(dialog, use_overlay=True)

    def _show_styled_warning_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        self._style_dialog(dialog, size="430x170")

        body = ttk.Frame(dialog, style="Dialog.TFrame")
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        ttk.Label(body, text=title, style="Dialog.TLabel", font=("Segoe UI Semibold", 12)).pack(anchor="w", pady=(0, 8))
        text_widget = tk.Text(
            body,
            height=3,
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            padx=2,
            pady=2,
            font=("Segoe UI", 10),
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", message)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(
            body,
            text="Понятно",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(anchor="e", pady=(10, 0))

        self._show_modal_dialog(dialog, use_overlay=True)

    def _show_styled_error_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        self._style_dialog(dialog, size="430x170")

        body = ttk.Frame(dialog, style="Dialog.TFrame")
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        ttk.Label(body, text=title, style="Dialog.TLabel", font=("Segoe UI Semibold", 12)).pack(anchor="w", pady=(0, 8))
        text_widget = tk.Text(
            body,
            height=3,
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            padx=2,
            pady=2,
            font=("Segoe UI", 10),
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", message)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(
            body,
            text="Понятно",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(anchor="e", pady=(10, 0))

        self._show_modal_dialog(dialog, use_overlay=True)

    def _ask_styled_confirm_dialog(self, title: str, message: str) -> bool:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        self._style_dialog(dialog, size="520x220")

        result = {"confirmed": False}
        body = ttk.Frame(dialog, style="Dialog.TFrame")
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)

        ttk.Label(body, text=title, style="Dialog.TLabel", font=("Segoe UI Semibold", 12)).pack(anchor="w", pady=(0, 8))
        text_widget = tk.Text(
            body,
            height=5,
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            padx=2,
            pady=2,
            font=("Segoe UI", 10),
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", message)
        text_widget.config(state=tk.DISABLED)

        button_row = ttk.Frame(body, style="Dialog.TFrame")
        button_row.pack(anchor="e", pady=(10, 0))

        def confirm() -> None:
            result["confirmed"] = True
            self._close_modal_dialog(dialog)

        ttk.Button(button_row, text="Да", style="Dialog.TButton", command=confirm).pack(side=tk.LEFT, padx=4)
        ttk.Button(
            button_row,
            text="Нет",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(side=tk.LEFT, padx=4)

        self._show_modal_dialog(dialog, use_overlay=True)
        self.wait_window(dialog)
        return bool(result["confirmed"])

    def _close_modal_dialog(self, dialog: tk.Toplevel) -> None:
        if self.modal_overlay and self.modal_overlay.winfo_exists():
            self.modal_overlay.destroy()
        self.modal_overlay = None
        if dialog.winfo_exists():
            dialog.destroy()

    def _show_modal_dialog(self, dialog: tk.Toplevel, *, use_overlay: bool = False) -> None:
        if use_overlay:
            self.modal_overlay = self._create_modal_overlay()
            self.modal_overlay.lift(self)
            self._animate_alpha(self.modal_overlay, target_alpha=0.25, duration_ms=150)
        self._center_dialog(dialog)
        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: self._close_modal_dialog(dialog))
        dialog.bind("<Escape>", lambda _event: self._close_modal_dialog(dialog))
        dialog.lift()
        dialog.focus_force()
        try:
            dialog.attributes("-alpha", 0.0)
        except tk.TclError:
            return
        self._animate_alpha(dialog, target_alpha=1.0, duration_ms=180)

    def clear_graphs_only(self) -> None:
        if self.fig_metrics and self.canvas_metrics:
            self.fig_metrics.clear()
            self.metrics_axes = None
            self.canvas_metrics.draw()

    def show_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise:
            self.clear_results()
            return

        try:
            self._validate_analysis_parameters()
        except ValueError as exc:
            self._show_styled_error_dialog("Ошибка ввода", str(exc))
            return

        if self.period_mode_var.get() == "Все данные":
            start_date = None
            end_date = None
        else:
            if not self.period_start_var.get().strip() or not self.period_end_var.get().strip():
                self.set_period_to_full_range(enterprise)

            if not self.period_start_var.get().strip() and not self.period_end_var.get().strip():
                self.current_enterprise = enterprise
                self.tree.delete(*self.tree.get_children())
                self.current_records = []
                self.selected_record_id = None
                self.clear_graphs_only()
                self.clear_statistics()
                return

            try:
                start_date = self.parse_date_string(self.period_start_var.get())
                end_date = self.parse_date_string(self.period_end_var.get())
                if not start_date or not end_date:
                    raise ValueError("Некорректный формат даты")
                if start_date > end_date:
                    raise ValueError("Начальная дата не может быть позже конечной")
                self._ensure_not_future_date(start_date, "\u041d\u0430\u0447\u0430\u043b\u044c\u043d\u0430\u044f \u0434\u0430\u0442\u0430 \u043f\u0435\u0440\u0438\u043e\u0434\u0430")
                self._ensure_not_future_date(end_date, "\u041a\u043e\u043d\u0435\u0447\u043d\u0430\u044f \u0434\u0430\u0442\u0430 \u043f\u0435\u0440\u0438\u043e\u0434\u0430")
            except ValueError as exc:
                self._show_styled_error_dialog(
                    "Ошибка даты",
                    "Некорректный формат даты.\n"
                    "Используйте ДД.ММ.ГГГГ или ГГГГ-ММ-ДД.\n\n"
                    f"Текст ошибки: {exc}",
                )
                self.clear_results()
                return

        self.current_enterprise = enterprise
        self.clear_statistics()
        self.current_records = self.repository.get_records(enterprise.id, start_date, end_date)
        self.tree.delete(*self.tree.get_children())

        for row_index, record in enumerate(self.current_records):
            self.tree.insert(
                "",
                tk.END,
                iid=str(record.id),
                tags=("even" if row_index % 2 == 0 else "odd",),
                values=(
                    record.period_date.strftime("%d.%m.%Y"),
                    f"{record.revenue:,.0f}",
                    f"{record.cost:,.0f}",
                    f"{record.fixed_expenses:,.0f}",
                    f"{record.variable_expenses:,.0f}",
                    f"{record.tax:,.0f}",
                    f"{record.net_profit:,.0f}",
                    f"{record.ros:.1f}",
                ),
            )

        self._sync_tree_selection()
        if self.current_records:
            self.plot_graphs()
        else:
            self.clear_graphs_only()
            self.clear_statistics()

    def _update_period_mode_ui(self) -> None:
        if self.period_mode_var.get() == "Период":
            if not self.period_from_label.winfo_manager():
                self.period_from_label.pack(side=tk.LEFT, padx=(0, 5))
            if not self.period_from_entry.winfo_manager():
                self.period_from_entry.pack(side=tk.LEFT, padx=5)
            if not self.period_to_label.winfo_manager():
                self.period_to_label.pack(side=tk.LEFT, padx=5)
            if not self.period_to_entry.winfo_manager():
                self.period_to_entry.pack(side=tk.LEFT, padx=5)
        else:
            self.period_from_label.pack_forget()
            self.period_from_entry.pack_forget()
            self.period_to_label.pack_forget()
            self.period_to_entry.pack_forget()

    def on_period_mode_change(self, _event) -> None:
        self._update_period_mode_ui()
        enterprise = self.get_selected_enterprise()
        if enterprise and self.period_mode_var.get() == "Период":
            self.set_period_to_full_range(enterprise)
        self.show_data()

    def plot_graphs(self, preserve_view: bool = False) -> None:
        if not self.current_records or not self.fig_metrics or not self.canvas_metrics:
            return

        current_xlim = None
        if preserve_view and self.metrics_axes:
            current_xlim = self.metrics_axes[0].get_xlim()

        x_indices = list(range(1, len(self.current_records) + 1))
        profits = [record.net_profit for record in self.current_records]
        ros_values = [record.ros for record in self.current_records]

        self.fig_metrics.clear()
        ax1 = self.fig_metrics.add_subplot(111)
        ax2 = ax1.twinx()
        self.metrics_axes = (ax1, ax2)

        ax1.plot(
            x_indices,
            profits,
            color="#1f5aa6",
            linewidth=2,
            marker="o",
            markersize=4,
            markerfacecolor="#1f5aa6",
            label="Чистая прибыль",
        )
        ax2.plot(
            x_indices,
            ros_values,
            color="#c0392b",
            linewidth=2,
            marker="s",
            markersize=4,
            markerfacecolor="#c0392b",
            label="ROS",
        )

        label_step = max(1, len(x_indices) // 12)
        for idx, x_val in enumerate(x_indices):
            if idx % label_step != 0 and idx != len(x_indices) - 1:
                continue
            ax1.annotate(
                f"{profits[idx] / 1000:.0f}k",
                (x_val, profits[idx]),
                xytext=(0, 6),
                textcoords="offset points",
                ha="center",
                fontsize=7,
                color="#4f647c",
                alpha=0.9,
            )
            ax2.annotate(
                f"{ros_values[idx]:.1f}%",
                (x_val, ros_values[idx]),
                xytext=(0, -10),
                textcoords="offset points",
                ha="center",
                fontsize=7,
                color="#7a2e28",
                alpha=0.9,
            )

        selected_index = self._get_selected_record_index()
        if selected_index is not None:
            selected_x = x_indices[selected_index]
            ax1.scatter(
                [selected_x],
                [profits[selected_index]],
                s=120,
                color="#0b3d91",
                edgecolors="white",
                linewidths=1.5,
                zorder=5,
            )
            ax2.scatter(
                [selected_x],
                [ros_values[selected_index]],
                s=120,
                color="#a61e11",
                edgecolors="white",
                linewidths=1.5,
                zorder=5,
            )
            ax1.axvline(selected_x, color="#7c93ad", linestyle=":", linewidth=1.2, alpha=0.7, zorder=0)

        ax1.set_xlabel("Период", fontsize=10, fontweight="bold")
        ax1.set_ylabel("Чистая прибыль, ₽", fontsize=10, fontweight="bold")
        ax2.set_ylabel("ROS, %", fontsize=10, fontweight="bold")
        ax1.grid(True, alpha=0.3)
        ax1.margins(x=0.04, y=0.18)
        ax2.margins(y=0.18)

        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _: f"{value / 1000:,.0f}K"))
        ax1.tick_params(axis="y", labelsize=8)
        ax2.tick_params(axis="y", labelsize=8)
        if current_xlim is None:
            x_min, x_max = 0.5, len(self.current_records) + 0.5
        else:
            data_min = 0.5
            data_max = len(self.current_records) + 0.5
            x_min = max(data_min, min(current_xlim[0], data_max))
            x_max = max(data_min, min(current_xlim[1], data_max))
            if x_max <= x_min:
                x_min, x_max = data_min, data_max
        ax1.set_xlim(x_min, x_max)
        ax2.set_xlim(x_min, x_max)
        self._update_graph_x_ticks()

        try:
            target_ros = float(self.target_ros_var.get())
            ax2.axhline(
                y=target_ros,
                color="#2e8b57",
                linestyle="--",
                linewidth=2,
                alpha=0.7,
                label=f"Целевой уровень: {target_ros}%",
            )
        except ValueError:
            pass

        profit_handle, profit_label = ax1.get_legend_handles_labels()
        ros_handle, ros_label = ax2.get_legend_handles_labels()
        ax1.legend(
            profit_handle + ros_handle,
            profit_label + ros_label,
            loc="upper center",
            bbox_to_anchor=(0.5, 0.985),
            ncol=3,
            fontsize=9,
            frameon=True,
            borderaxespad=0.3,
            prop={"weight": "bold"},
        )

        self.fig_metrics.subplots_adjust(top=0.9, left=0.1, right=0.9, bottom=0.16)
        self.canvas_metrics.draw()

    def _get_selected_record_index(self) -> int | None:
        if self.selected_record_id is None:
            return None
        for index, record in enumerate(self.current_records):
            if record.id == self.selected_record_id:
                return index
        return None

    def _update_graph_x_ticks(self) -> None:
        if not self.metrics_axes or not self.current_records:
            return

        ax1, ax2 = self.metrics_axes
        x_min, x_max = ax1.get_xlim()
        visible_start = max(1, int(round(x_min)))
        visible_end = min(len(self.current_records), int(round(x_max)))
        visible_count = max(1, visible_end - visible_start + 1)
        tick_mode = self.graph_tick_mode_var.get()
        if tick_mode == "Каждый месяц":
            tick_step = 1
        elif tick_mode == "Через 2 месяца":
            tick_step = 2
        else:
            tick_step = max(1, visible_count // 10)

        tick_positions = list(range(visible_start, visible_end + 1, tick_step))
        if not tick_positions or tick_positions[-1] != visible_end:
            tick_positions.append(visible_end)
        tick_labels = [self.current_records[position - 1].period_date.strftime("%m.%y") for position in tick_positions]
        ax1.set_xticks(tick_positions)
        ax1.set_xticklabels(tick_labels, rotation=35, ha="right", fontsize=8)
        ax2.set_xticks(tick_positions)
        self.fig_metrics.tight_layout(rect=(0.02, 0.03, 0.98, 0.94))
        self.fig_metrics.subplots_adjust(top=0.9)

    def _sync_tree_selection(self) -> None:
        selected_iid = str(self.selected_record_id) if self.selected_record_id is not None else ""
        current_items = set(self.tree.get_children())
        self.suppress_tree_event = True
        try:
            if selected_iid and selected_iid in current_items:
                self.tree.selection_set(selected_iid)
                self.tree.focus(selected_iid)
                self.tree.see(selected_iid)
            else:
                self.tree.selection_remove(*self.tree.selection())
        finally:
            self.suppress_tree_event = False

    def _select_record_by_index(self, index: int) -> None:
        if index < 0 or index >= len(self.current_records):
            return
        self.selected_record_id = self.current_records[index].id
        self._sync_tree_selection()
        self.plot_graphs(preserve_view=True)

    def test_hypothesis(self) -> None:
        if len(self.current_records) < 2:
            self._show_styled_warning_dialog("Ошибка", "Недостаточно данных для проверки гипотезы")
            return

        try:
            target_ros, alpha = self._validate_analysis_parameters()
        except ValueError as exc:
            self._show_styled_error_dialog("Ошибка ввода", str(exc))
            return

        ros_values = [record.ros for record in self.current_records]
        try:
            t_stat, p_value = t_test_one_sample(ros_values, target_ros)
            avg_ros, std_ros = summarize_ros(ros_values)
        except ValueError as exc:
            self._show_styled_error_dialog("Ошибка вычислений", str(exc))
            return

        verdict = "Не отклоняется" if p_value >= alpha else "Отклоняется"
        recommendation = (
            "Рекомендуется к инвестированию"
            if p_value >= alpha
            else "Не рекомендуется к инвестированию"
        )
        quick_verdict = "рекомендуется к инвестированию" if p_value >= alpha else "не рекомендуется к инвестированию"
        self.quick_recommendation_var.set(f"Практический вывод: {quick_verdict}")
        profits = [record.net_profit for record in self.current_records]
        avg_profit = sum(profits) / len(profits)
        min_ros = min(ros_values)
        max_ros = max(ros_values)
        hypothesis_text = (
            f"H0: средний ROS не ниже целевого уровня {target_ros:.1f}%.\n"
            f"H1: средний ROS ниже целевого уровня {target_ros:.1f}%."
        )
        verdict_explanation = (
            "p-уровень не меньше α, поэтому статистически недостаточно оснований "
            "считать, что средний ROS ниже целевого."
            if p_value >= alpha
            else "p-уровень меньше α, поэтому есть статистические основания "
            "считать, что средний ROS ниже целевого."
        )

        result_lines = [
            "Статистические показатели",
            f"Периодов в расчёте: {len(self.current_records)}",
            f"Период анализа: {self.current_records[0].period_date.strftime('%d.%m.%Y')} - "
            f"{self.current_records[-1].period_date.strftime('%d.%m.%Y')}",
            "",
            "Что было рассчитано:",
            f"Средняя чистая прибыль: {avg_profit:,.0f} ₽",
            f"Средняя ROS: {avg_ros:.1f}% "
            f"(ROS = чистая прибыль / выручка * 100)",
            f"Минимальный ROS: {min_ros:.1f}%",
            f"Максимальный ROS: {max_ros:.1f}%",
            f"Стандартное отклонение ROS: {std_ros:.1f}% "
            f"(показывает, насколько ROS колебался от периода к периоду)",
            "",
            "Проверка гипотезы:",
            hypothesis_text,
            f"t-статистика: {t_stat:.2f}",
            f"p-уровень: {p_value:.3f}",
            f"Уровень значимости α: {alpha:.2f}",
            f"Вердикт: {verdict}",
            verdict_explanation,
            "",
            f"Практический вывод: {recommendation}",
        ]
        self.render_statistics(result_lines)

        self.analysis_result = {
            "avg_ros": avg_ros,
            "std_ros": std_ros,
            "min_ros": min_ros,
            "max_ros": max_ros,
            "avg_profit": avg_profit,
            "t_stat": t_stat,
            "p_value": p_value,
            "verdict": verdict,
            "recommendation": recommendation,
            "target_ros": target_ros,
            "alpha": alpha,
            "enterprise": self.enterprise_var.get(),
            "enterprise_id": self.current_enterprise.id if self.current_enterprise else None,
            "date_created": datetime.now().strftime("%d.%m.%Y"),
            "period_start": self.current_records[0].period_date.strftime("%Y-%m-%d"),
            "period_end": self.current_records[-1].period_date.strftime("%Y-%m-%d"),
        }
        try:
            self.current_report = self._save_current_report_to_database()
        except psycopg.Error as exc:
            self.current_report = None
            self._show_styled_warning_dialog(
                "Автосохранение отчёта",
                "Показатели рассчитаны, но отчёт не удалось сохранить в базу данных.\n\n"
                f"Текст ошибки: {exc}",
            )
            return
        if self.current_report:
            self.analysis_result["report_id"] = self.current_report.id

    def show_saved_reports(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Сохранённые отчёты")
        self._style_dialog(dialog, size="980x620", resizable=(True, True))

        sort_var = tk.StringVar(value="Дата (новые сначала)")
        search_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        search_frame.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=42)
        search_entry.pack(side=tk.LEFT, padx=6)
        ttk.Label(search_frame, text="Сортировка:").pack(side=tk.LEFT, padx=(10, 6))
        sort_combo = ttk.Combobox(
            search_frame,
            textvariable=sort_var,
            state="readonly",
            width=30,
            values=(
                "Дата (новые сначала)",
                "Дата (старые сначала)",
                "Предприятие (А-Я)",
                "Предприятие (Я-А)",
            ),
        )
        sort_combo.pack(side=tk.LEFT, padx=4)
        self._setup_combobox_behavior(sort_combo)

        columns = ("enterprise", "name", "date_created", "period_start", "period_end")
        tree_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)
        headings = {
            "enterprise": "Предприятие",
            "name": "Название",
            "date_created": "Дата формирования",
            "period_start": "Период с",
            "period_end": "Период по",
        }
        widths = {
            "enterprise": 220,
            "name": 280,
            "date_created": 140,
            "period_start": 120,
            "period_end": 120,
        }
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=widths[column], anchor=tk.W)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, style="Vertical.TScrollbar", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        rename_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        rename_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        ttk.Label(rename_frame, text="Название отчёта:").pack(side=tk.LEFT)
        report_name_var = tk.StringVar()
        ttk.Entry(rename_frame, textvariable=report_name_var, width=52).pack(side=tk.LEFT, padx=6)

        report_by_id: dict[int, FinancialReport] = {}

        def get_sorted_reports(reports: list[FinancialReport], enterprise_names: dict[int, str]) -> list[FinancialReport]:
            mode = sort_var.get()
            if mode == "Дата (старые сначала)":
                return sorted(reports, key=lambda item: (item.date_created, item.id or 0))
            if mode == "Предприятие (А-Я)":
                return sorted(
                    reports,
                    key=lambda item: (enterprise_names.get(item.enterprise_id, ""), item.date_created, item.id or 0),
                )
            if mode == "Предприятие (Я-А)":
                return sorted(
                    reports,
                    key=lambda item: (enterprise_names.get(item.enterprise_id, ""), item.date_created, item.id or 0),
                    reverse=True,
                )
            return sorted(reports, key=lambda item: (item.date_created, item.id or 0), reverse=True)

        def load_reports(*_args) -> None:
            nonlocal report_by_id
            tree.delete(*tree.get_children())

            enterprise_names = {item.id: item.name for item in self.repository.list_enterprises()}
            reports = self.repository.list_financial_reports()
            reports = get_sorted_reports(reports, enterprise_names)

            query = search_var.get().strip().lower()
            if query:
                def matches(report: FinancialReport) -> bool:
                    enterprise_title = enterprise_names.get(report.enterprise_id, "")
                    searchable = [
                        enterprise_title,
                        report.name,
                        report.date_created.strftime("%d.%m.%Y"),
                        report.period_start.strftime("%d.%m.%Y"),
                        report.period_end.strftime("%d.%m.%Y"),
                    ]
                    return any(query in item.lower() for item in searchable)

                reports = [report for report in reports if matches(report)]

            report_by_id = {int(report.id): report for report in reports if report.id is not None}

            for report in reports:
                if report.id is None:
                    continue
                tree.insert(
                    "",
                    tk.END,
                    iid=str(report.id),
                    values=(
                        enterprise_names.get(report.enterprise_id, "Неизвестное предприятие"),
                        report.name,
                        report.date_created.strftime("%d.%m.%Y"),
                        report.period_start.strftime("%d.%m.%Y"),
                        report.period_end.strftime("%d.%m.%Y"),
                    ),
                )

        def clear_search() -> None:
            search_var.set("")
            load_reports()
            search_entry.focus_set()

        ttk.Button(search_frame, text="Сбросить", style="Dialog.TButton", command=clear_search).pack(side=tk.LEFT, padx=4)

        def get_selected_report_id() -> int | None:
            selection = tree.selection()
            return int(selection[0]) if selection else None

        def on_tree_select(_event=None) -> None:
            report_id = get_selected_report_id()
            report = report_by_id.get(report_id) if report_id is not None else None
            report_name_var.set(report.name if report else "")

        def rename_report() -> None:
            report_id = get_selected_report_id()
            if report_id is None:
                self._show_styled_warning_dialog("Ошибка", "Выберите отчёт")
                return

            new_name = report_name_var.get().strip()
            if not new_name:
                self._show_styled_warning_dialog("Ошибка", "Введите название отчёта")
                return

            updated = self.repository.update_financial_report_name(report_id, new_name)
            if updated is None:
                self._show_styled_error_dialog("Ошибка", "Отчёт не найден")
                return

            load_reports()
            tree.selection_set(str(report_id))
            tree.focus(str(report_id))
            report_name_var.set(updated.name)

        def open_report() -> None:
            report_id = get_selected_report_id()
            if report_id is None:
                self._show_styled_warning_dialog("Ошибка", "Выберите отчёт")
                return
            self._load_saved_report(report_id)
            self._close_modal_dialog(dialog)

        def export_report() -> None:
            report_id = get_selected_report_id()
            if report_id is None:
                self._show_styled_warning_dialog("Ошибка", "Выберите отчёт")
                return
            report = self.repository.get_financial_report(report_id)
            if report is None:
                self._show_styled_error_dialog("Ошибка", "Отчёт не найден")
                return
            metric_values = self.repository.get_financial_report_metric_values(report_id)
            records = self.repository.get_financial_report_records(report_id)
            file_path = filedialog.asksaveasfilename(
                defaultextension=".html",
                filetypes=[("HTML файл", "*.html"), ("Текстовый файл", "*.txt")],
                title="Экспорт сохранённого отчёта",
            )
            if not file_path:
                return
            report_html = self._build_analysis_report_html(
                analysis_result=self._analysis_result_from_report(report, metric_values),
                records=records,
            )
            with open(file_path, "w", encoding="utf-8") as report_file:
                report_file.write(report_html)
            self._show_styled_info_dialog("Успех", f"Отчёт сохранён:\n{file_path}")

        def delete_report() -> None:
            report_id = get_selected_report_id()
            if report_id is None:
                self._show_styled_warning_dialog("Ошибка", "Выберите отчёт")
                return
            confirmed = self._ask_styled_confirm_dialog(
                "Удалить отчёт",
                "Удалить выбранный отчёт из базы данных?",
            )
            if not confirmed:
                return
            self.repository.delete_financial_report(report_id)
            report_name_var.set("")
            load_reports()

        button_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(button_frame, text="Открыть", style="Dialog.TButton", command=open_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Переименовать", style="Dialog.TButton", command=rename_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Экспорт HTML", style="Dialog.TButton", command=export_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Удалить", style="Dialog.TButton", command=delete_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(
            button_frame,
            text="Закрыть",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(side=tk.RIGHT, padx=4)

        tree.bind("<<TreeviewSelect>>", on_tree_select)
        sort_combo.bind("<<ComboboxSelected>>", load_reports)
        search_entry.bind("<KeyRelease>", load_reports)

        load_reports()
        self._show_modal_dialog(dialog, use_overlay=True)

    def save_report(self) -> None:
        if not self.analysis_result:
            self._show_styled_warning_dialog("Ошибка", "Сначала проведите проверку гипотезы")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
                filetypes=[("HTML файл", "*.html"), ("Текстовый файл", "*.txt")],
            title="Сохранить аналитический отчёт",
        )
        if not file_path:
            return

        report_html = self._build_analysis_report_html()
        with open(file_path, "w", encoding="utf-8") as report_file:
            report_file.write(report_html)

        if self.current_enterprise and self.current_records:
            self.repository.save_financial_report(
                FinancialReport(
                    id=None,
                    enterprise_id=self.current_enterprise.id,
                    name=f"Аналитический отчет {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                    date_created=datetime.now().date(),
                    period_start=self.current_records[0].period_date,
                    period_end=self.current_records[-1].period_date,
                ),
                {
                    "target_ros": float(self.analysis_result["target_ros"]),
                    "alpha": float(self.analysis_result["alpha"]),
                    "avg_ros": float(self.analysis_result["avg_ros"]),
                    "std_ros": float(self.analysis_result["std_ros"]),
                    "min_ros": float(self.analysis_result["min_ros"]),
                    "max_ros": float(self.analysis_result["max_ros"]),
                    "avg_profit": float(self.analysis_result["avg_profit"]),
                    "t_stat": float(self.analysis_result["t_stat"]),
                    "p_value": float(self.analysis_result["p_value"]),
                },
            )

        self._show_styled_info_dialog("Успех", f"Отчёт сохранён:\n{file_path}")

    def _load_saved_report(self, report_id: int) -> None:
        report = self.repository.get_financial_report(report_id)
        if report is None:
            self._show_styled_error_dialog("Ошибка", "Отчёт не найден")
            return

        metric_values = self.repository.get_financial_report_metric_values(report_id)
        analysis_result = self._analysis_result_from_report(report, metric_values)
        enterprise = next((item for item in self.repository.list_enterprises() if item.id == report.enterprise_id), None)
        if enterprise is None:
            self._show_styled_error_dialog("Ошибка", "Предприятие для отчёта не найдено")
            return

        self.enterprise_var.set(enterprise.name)
        self.period_mode_var.set("Период")
        self._update_period_mode_ui()
        self.period_start_var.set(report.period_start.strftime("%Y-%m-%d"))
        self.period_end_var.set(report.period_end.strftime("%Y-%m-%d"))
        self.show_data()
        self.analysis_result = analysis_result
        self.render_statistics(self._statistics_lines_from_analysis_result(analysis_result))
        recommendation_text = str(analysis_result.get("recommendation", "")).lower()
        quick_verdict = (
            "не рекомендуется к инвестированию"
            if recommendation_text.startswith("не ")
            else "рекомендуется к инвестированию"
        )
        self.quick_recommendation_var.set(f"Практический вывод: {quick_verdict}")

    def _validate_analysis_parameters(self) -> tuple[float, float]:
        target_ros = float(self.target_ros_var.get())
        alpha = float(self.alpha_var.get())
        if not 0 < alpha < 1:
            raise ValueError("α должен быть в диапазоне (0, 1)")
        if target_ros < 0:
            raise ValueError("Целевой ROS не может быть отрицательным")
        return target_ros, alpha

    def _analysis_result_from_report(self, report: FinancialReport, metric_values: dict[str, float]) -> dict:
        alpha = float(metric_values.get("alpha", 0.05))
        p_value = float(metric_values.get("p_value", 1.0))
        verdict = "Не отклоняется" if p_value >= alpha else "Отклоняется"
        recommendation = (
            "Рекомендуется к инвестированию"
            if p_value >= alpha
            else "Не рекомендуется к инвестированию"
        )
        enterprise = next((item for item in self.repository.list_enterprises() if item.id == report.enterprise_id), None)
        return {
            "avg_ros": float(metric_values.get("avg_ros", 0.0)),
            "std_ros": float(metric_values.get("std_ros", 0.0)),
            "min_ros": float(metric_values.get("min_ros", 0.0)),
            "max_ros": float(metric_values.get("max_ros", 0.0)),
            "avg_profit": float(metric_values.get("avg_profit", 0.0)),
            "t_stat": float(metric_values.get("t_stat", 0.0)),
            "p_value": p_value,
            "verdict": verdict,
            "recommendation": recommendation,
            "target_ros": float(metric_values.get("target_ros", 0.0)),
            "alpha": alpha,
            "enterprise": enterprise.name if enterprise else "Неизвестное предприятие",
            "enterprise_id": report.enterprise_id,
            "date_created": report.date_created.strftime("%d.%m.%Y"),
            "period_start": report.period_start.strftime("%Y-%m-%d"),
            "period_end": report.period_end.strftime("%Y-%m-%d"),
        }

    def _statistics_lines_from_analysis_result(self, analysis_result: dict) -> list[str]:
        hypothesis_text = (
            f"H0: средний ROS не ниже целевого уровня {analysis_result['target_ros']:.1f}%.\n"
            f"H1: средний ROS ниже целевого уровня {analysis_result['target_ros']:.1f}%."
        )
        verdict_explanation = (
            "p-уровень не меньше α, поэтому статистически недостаточно оснований считать, что средний ROS ниже целевого."
            if analysis_result["p_value"] >= analysis_result["alpha"]
            else "p-уровень меньше α, поэтому есть статистические основания считать, что средний ROS ниже целевого."
        )
        return [
            "Статистические показатели",
            f"Период анализа: {analysis_result['period_start']} - {analysis_result['period_end']}",
            "",
            "Что было рассчитано:",
            f"Средняя чистая прибыль: {analysis_result['avg_profit']:,.0f} ₽",
            f"Средняя ROS: {analysis_result['avg_ros']:.1f}%",
            f"Минимальный ROS: {analysis_result['min_ros']:.1f}%",
            f"Максимальный ROS: {analysis_result['max_ros']:.1f}%",
            f"Стандартное отклонение ROS: {analysis_result['std_ros']:.1f}%",
            "",
            "Проверка гипотезы:",
            hypothesis_text,
            f"t-статистика: {analysis_result['t_stat']:.2f}",
            f"p-уровень: {analysis_result['p_value']:.3f}",
            f"Уровень значимости α: {analysis_result['alpha']:.2f}",
            f"Вердикт: {analysis_result['verdict']}",
            verdict_explanation,
            "",
            f"Практический вывод: {analysis_result['recommendation']}",
        ]

    def _build_analysis_report_html(self, analysis_result: dict | None = None, records: list[FinancialRecord] | None = None) -> str:
        if analysis_result is None:
            analysis_result = self.analysis_result
        if records is None:
            records = self.current_records

        def esc(value: object) -> str:
            return html.escape(str(value))

        summary_rows = [
            ("Предприятие", analysis_result["enterprise"]),
            ("Дата анализа", analysis_result["date_created"]),
            ("Период анализа", f'{analysis_result["period_start"]} - {analysis_result["period_end"]}'),
            ("Целевой уровень ROS", f'{analysis_result["target_ros"]:.1f}%'),
            ("Уровень значимости α", f'{analysis_result["alpha"]:.2f}'),
            ("Средняя чистая прибыль", f'{analysis_result["avg_profit"]:,.0f} ₽'),
            ("Средний ROS", f'{analysis_result["avg_ros"]:.1f}%'),
            ("Минимальный ROS", f'{analysis_result["min_ros"]:.1f}%'),
            ("Максимальный ROS", f'{analysis_result["max_ros"]:.1f}%'),
            ("Стандартное отклонение ROS", f'{analysis_result["std_ros"]:.1f}%'),
            ("t-статистика", f'{analysis_result["t_stat"]:.2f}'),
            ("p-уровень", f'{analysis_result["p_value"]:.3f}'),
            ("Вердикт гипотезы", analysis_result["verdict"]),
        ]

        summary_html = "".join(f"<tr><th>{esc(label)}</th><td>{esc(value)}</td></tr>" for label, value in summary_rows)

        data_rows_html = "".join(
            (
                "<tr>"
                f"<td>{record.period_date.strftime('%d.%m.%Y')}</td>"
                f"<td>{record.revenue:,.0f}</td>"
                f"<td>{record.cost:,.0f}</td>"
                f"<td>{record.fixed_expenses:,.0f}</td>"
                f"<td>{record.variable_expenses:,.0f}</td>"
                f"<td>{record.tax:,.0f}</td>"
                f"<td>{record.net_profit:,.0f}</td>"
                f"<td>{record.ros:.1f}</td>"
                "</tr>"
            )
            for record in records
        )
        table_colgroup_html = """
            <colgroup>
              <col style="width: 12%;">
              <col style="width: 14%;">
              <col style="width: 14%;">
              <col style="width: 14%;">
              <col style="width: 14%;">
              <col style="width: 12%;">
              <col style="width: 16%;">
              <col style="width: 8%;">
            </colgroup>
        """

        return f"""<!DOCTYPE html>
<html lang=\"ru\">
<head>
  <meta charset=\"UTF-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
  <title>Аналитический отчёт</title>
  <style>
    :root {{
      --bg-1: #f4f8fd;
      --bg-2: #eaf2fb;
      --surface: #ffffff;
      --surface-soft: #f6fafe;
      --text: #1e3550;
      --muted: #5f7792;
      --primary-soft: #dceaf9;
      --border: #d3e1ef;
      --success: #0f8a5f;
      --success-bg: #e6f6ef;
      --danger: #ba3b3b;
      --danger-bg: #fdeeee;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: \"Segoe UI\", \"Arial\", sans-serif;
      color: var(--text);
      background: linear-gradient(165deg, var(--bg-1), var(--bg-2));
      padding: 28px 18px 36px;
    }}
    .page {{
      max-width: 1160px;
      margin: 0 auto;
    }}
    .grid {{
      display: grid;
      grid-template-columns: 1fr;
      gap: 14px;
    }}
    .card {{
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 14px;
      box-shadow: 0 6px 18px rgba(42, 86, 132, 0.08);
      overflow: hidden;
    }}
    .card-head {{
      background: var(--surface-soft);
      padding: 12px 16px;
      border-bottom: 1px solid var(--border);
      font-size: 16px;
      font-weight: 700;
      color: #1d4269;
    }}
    .card-body {{
      padding: 14px 16px 16px;
    }}
    table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
    }}
    th, td {{
      padding: 10px 12px;
      text-align: left;
      border-bottom: 1px solid #e9f0f7;
      font-size: 14px;
    }}
    .meta th {{
      width: 42%;
      color: #355b82;
      font-weight: 600;
      background: #fbfdff;
    }}
    .meta td {{
      color: var(--text);
      font-weight: 500;
    }}
    .period-table-header {{
      border: 1px solid #d4e3f3;
      border-bottom: none;
      border-radius: 10px 10px 0 0;
      overflow: hidden;
      background: var(--primary-soft);
    }}
    .period-table-header .period-table thead th {{
      background: var(--primary-soft);
      color: #17426e;
      font-weight: 700;
      white-space: nowrap;
      border-bottom: 1px solid #c9ddf1;
    }}
    .period-table-body .period-table tbody td {{
      white-space: nowrap;
      background: #ffffff;
    }}
    .period-table-body .period-table tbody td:nth-child(n+2),
    .period-table-header .period-table thead th:nth-child(n+2) {{
      text-align: right;
      font-variant-numeric: tabular-nums;
    }}
    .period-table-body .period-table tbody tr:nth-child(even) {{
      background: #fbfdff;
    }}
    .period-table-body .period-table tbody tr:hover {{
      background: #f2f8ff;
    }}
    .period-table-body {{
      max-height: 460px;
      overflow: auto;
      border: 1px solid #d4e3f3;
      border-top: none;
      border-radius: 0 0 10px 10px;
      background: #ffffff;
    }}
    .table-wrap {{
      overflow: hidden;
    }}
    .period-table {{
      min-width: 980px;
      border-collapse: collapse;
      table-layout: fixed;
    }}
    .footer-note {{
      margin-top: 14px;
      color: var(--muted);
      font-size: 13px;
      text-align: right;
    }}
    @media (max-width: 860px) {{
      body {{
        padding: 16px 10px 24px;
      }}
      th, td {{
        font-size: 13px;
        padding: 8px 10px;
      }}
    }}
  </style>
</head>
<body>
  <div class=\"page\">
    <div class=\"grid\">
      <section class=\"card\">
        <div class=\"card-head\">Рассчитанные показатели</div>
        <div class=\"card-body\">
          <table class=\"meta\">
            {summary_html}
          </table>
        </div>
      </section>

      <section class=\"card\">
        <div class=\"card-head\">Данные по периодам</div>
        <div class=\"card-body table-wrap\">
          <div class=\"period-table-header\">
            <table class=\"period-table\">
              {table_colgroup_html}
              <thead>
                <tr>
                  <th>Дата</th>
                  <th>Выручка, ₽</th>
                  <th>Себестоимость, ₽</th>
                  <th>Пост. издержки, ₽</th>
                  <th>Перем. издержки, ₽</th>
                  <th>Налог, ₽</th>
                  <th>Чистая прибыль, ₽</th>
                  <th>ROS, %</th>
                </tr>
              </thead>
            </table>
          </div>
          <div class=\"period-table-body\">
            <table class=\"period-table\">
              {table_colgroup_html}
              <tbody>
                {data_rows_html}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    </div>
    <div class=\"footer-note\">Сформировано в приложении «{esc(APP_TITLE)}»</div>
  </div>
</body>
</html>
"""
    def add_enterprise(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Добавить предприятие")
        self._style_dialog(dialog, size="360x150")

        ttk.Label(dialog, text="Название предприятия:", style="Dialog.TLabel").pack(pady=8)
        name_entry = ttk.Entry(dialog, width=34)
        name_entry.pack(pady=5)

        def save() -> None:
            name = name_entry.get().strip()
            if not name:
                self._show_styled_warning_dialog("Ошибка", "Введите название предприятия")
                return
            if name in self.enterprise_by_name:
                self._show_styled_error_dialog("Ошибка", "Предприятие с таким названием уже существует")
                return
            try:
                enterprise = self.repository.add_enterprise(name)
            except psycopg.errors.UniqueViolation:
                self._show_styled_error_dialog("Ошибка", "Предприятие с таким названием уже существует")
                return
            self.enterprise_by_name[enterprise.name] = enterprise
            self.update_enterprise_list()
            self.enterprise_var.set(enterprise.name)
            self.clear_statistics()
            self._close_modal_dialog(dialog)
            self.refresh_current_view(rerun_analysis=False)

        button_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        button_frame.pack(pady=12)
        ttk.Button(button_frame, text="Сохранить", style="Dialog.TButton", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Отмена",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(side=tk.LEFT, padx=5)
        self._show_modal_dialog(dialog, use_overlay=True)
        name_entry.focus_set()

    def add_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if enterprise:
            self.open_record_dialog(enterprise)

    def delete_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise or self.selected_record_id is None:
            self._show_styled_warning_dialog("Ошибка", "Выберите строку для удаления")
            return

        record = next((item for item in self.current_records if item.id == self.selected_record_id), None)
        if not record:
            self._show_styled_error_dialog("Ошибка", "Данные не найдены")
            return

        confirmed = self._ask_styled_confirm_dialog(
            "Удалить данные",
            f"Удалить данные за {record.period_date.strftime('%d.%m.%Y')}?\nУдаление будет выполнено и из базы данных.",
        )
        if not confirmed:
            return

        try:
            self.repository.delete_record(record.id)
        except ValueError as exc:
            self._show_styled_error_dialog("Ошибка", str(exc))
            return
        except psycopg.Error as exc:
            self._show_styled_error_dialog("Ошибка PostgreSQL", str(exc))
            return

        self.selected_record_id = None
        self.clear_statistics()
        self.refresh_current_view(rerun_analysis=False)
        self._show_styled_info_dialog("Успех", "Данные успешно удалены")

    def delete_enterprise(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise:
            return

        confirmed = self._ask_styled_confirm_dialog(
            "Удалить предприятие",
            f"Удалить предприятие «{enterprise.name}»?\nВсе его показатели тоже будут удалены.",
        )
        if not confirmed:
            return

        self.repository.delete_enterprise(enterprise.id)
        self.update_enterprise_list()
        if self.enterprise_var.get() == enterprise.name:
            names = list(self.enterprise_by_name.keys())
            self.enterprise_var.set(names[0] if names else "")
        self.selected_record_id = None
        self.current_enterprise = None
        self.current_records = []
        self.clear_statistics()
        self.tree.delete(*self.tree.get_children())
        self.clear_graphs_only()
        if self.enterprise_var.get():
            self.refresh_current_view(rerun_analysis=False)

    def import_data(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Импорт данных",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return

        enterprise = self.get_selected_enterprise()
        if not enterprise:
            return

        try:
            self.config(cursor="watch")
            self.update_idletasks()
            dataframe = self._read_import_dataframe(file_path)
            dataframe = self._normalize_import_dataframe(dataframe, enterprise)
            overlapping_dates = self._find_existing_import_dates(dataframe, enterprise)
            if overlapping_dates:
                overlap_preview = self._format_import_date_ranges(overlapping_dates)
                confirmed = self._ask_styled_confirm_dialog(
                    "Подтвердите перезапись",
                    f"Будут перезаписаны уже существующие месяцы: {len(overlapping_dates)} шт.\n\n"
                    "Периоды:\n"
                    f"{overlap_preview}\n\n"
                    "Продолжить импорт?",
                )
                if not confirmed:
                    return
            imported_enterprises, imported_dates = self._save_imported_dataframe(dataframe, enterprise)
        except ValueError as exc:
            self._show_styled_error_dialog("Ошибка импорта", str(exc))
            return
        except psycopg.Error as exc:
            self._show_styled_error_dialog("Ошибка PostgreSQL", str(exc))
            return
        except Exception as exc:
            self._show_styled_error_dialog("Ошибка импорта", f"Не удалось импортировать файл:\n{exc}")
            return
        finally:
            self.config(cursor="")

        self.update_enterprise_list()
        self.enterprise_var.set(enterprise.name)
        enterprise = self.enterprise_by_name.get(enterprise.name)
        if enterprise:
            self.set_period_to_full_range(enterprise)

        if imported_dates:
            self.period_start_var.set(min(imported_dates).strftime("%Y-%m-%d"))
            self.period_end_var.set(max(imported_dates).strftime("%Y-%m-%d"))

        self.clear_statistics()
        self.show_data()
        self._show_styled_info_dialog(
            "Импорт завершён",
            f"Импортировано строк: {len(imported_dates)}\n"
            f"Предприятие: {enterprise.name}",
        )

    def edit_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise or self.selected_record_id is None:
            self._show_styled_warning_dialog("Ошибка", "Выберите строку для редактирования")
            return

        record = next((item for item in self.current_records if item.id == self.selected_record_id), None)
        if not record:
            self._show_styled_error_dialog("Ошибка", "Данные не найдены")
            return
        self.open_record_dialog(enterprise, record)

    def _read_import_dataframe(self, file_path: str) -> pd.DataFrame:
        file_path_lower = file_path.lower()
        if file_path_lower.endswith(".csv"):
            for encoding in ("utf-8-sig", "utf-8", "cp1251"):
                try:
                    return pd.read_csv(file_path, encoding=encoding)
                except UnicodeDecodeError:
                    continue
            raise ValueError("Не удалось прочитать CSV. Проверьте кодировку файла.")
        if file_path_lower.endswith((".xlsx", ".xls")):
            return pd.read_excel(file_path)
        raise ValueError("Поддерживаются только файлы CSV и Excel")

    def _normalize_import_dataframe(self, dataframe: pd.DataFrame, enterprise: Enterprise) -> pd.DataFrame:
        if dataframe.empty:
            raise ValueError("Файл импорта пуст")

        normalized_columns: dict[str, str] = {}
        for column in dataframe.columns:
            normalized_key = self._normalize_import_column_name(column)
            normalized_columns[column] = IMPORT_COLUMN_ALIASES.get(normalized_key, normalized_key)

        dataframe = dataframe.rename(columns=normalized_columns)
        required_columns = {
            "period_date",
            "revenue",
            "cost",
            "fixed_expenses",
            "variable_expenses",
            "tax",
        }
        missing_columns = required_columns - set(dataframe.columns)
        if missing_columns:
            missing_list = ", ".join(sorted(missing_columns))
            raise ValueError(f"В файле не хватает обязательных колонок: {missing_list}")

        dataframe["enterprise_name"] = enterprise.name

        period_dates = dataframe["period_date"].apply(self._parse_import_period_date)
        invalid_rows = dataframe.index[period_dates.isna()].tolist()
        if invalid_rows:
            raise ValueError(f"Ошибка в строке {invalid_rows[0] + 2}: Некорректная дата")
        today = date.today()
        future_rows = dataframe.index[period_dates.apply(lambda value: value > today)].tolist()
        if future_rows:
            raise ValueError(f"\u041e\u0448\u0438\u0431\u043a\u0430 \u0432 \u0441\u0442\u0440\u043e\u043a\u0435 {future_rows[0] + 2}: \u0434\u0430\u0442\u0430 \u043d\u0435 \u043c\u043e\u0436\u0435\u0442 \u0431\u044b\u0442\u044c \u0432 \u0431\u0443\u0434\u0443\u0449\u0435\u043c")

        dataframe["period_date"] = period_dates

        return dataframe

    @staticmethod
    def _normalize_import_column_name(column: object) -> str:
        normalized_key = str(column).strip().lower().replace("ё", "е")
        normalized_key = normalized_key.replace("₽", "").replace("в‚₽", "").replace("%", "")
        normalized_key = re.sub(r"[\s/\\\-.,:;()]+", "_", normalized_key)
        normalized_key = re.sub(r"_+", "_", normalized_key).strip("_")
        return normalized_key

    @staticmethod
    def _parse_import_period_date(value: object) -> date | None:
        if pd.isna(value):
            return None
        if isinstance(value, pd.Timestamp):
            return value.date()
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value

        raw_value = str(value).strip()
        if not raw_value:
            return None

        for date_format in ("%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(raw_value, date_format).date()
            except ValueError:
                continue

        parsed = pd.to_datetime(raw_value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.date()

    def _find_existing_import_dates(self, dataframe: pd.DataFrame, enterprise: Enterprise) -> list[date]:
        imported_dates = sorted({period_date for period_date in dataframe["period_date"].tolist() if period_date})
        if not imported_dates:
            return []

        existing_records = self.repository.get_records(
            enterprise.id,
            start_date=min(imported_dates),
            end_date=max(imported_dates),
        )
        existing_dates = {record.period_date for record in existing_records}
        return [period_date for period_date in imported_dates if period_date in existing_dates]

    @staticmethod
    def _format_import_date_ranges(dates: list[date]) -> str:
        if not dates:
            return ""

        sorted_dates = sorted(dates)
        ranges: list[tuple[date, date]] = []
        range_start = sorted_dates[0]
        range_end = sorted_dates[0]

        for current_date in sorted_dates[1:]:
            expected_year = range_end.year + (1 if range_end.month == 12 else 0)
            expected_month = 1 if range_end.month == 12 else range_end.month + 1
            if current_date.year == expected_year and current_date.month == expected_month:
                range_end = current_date
                continue
            ranges.append((range_start, range_end))
            range_start = current_date
            range_end = current_date

        ranges.append((range_start, range_end))
        return "; ".join(
            start.strftime("%m.%Y") if start == end else f"{start.strftime('%m.%Y')}-{end.strftime('%m.%Y')}"
            for start, end in ranges
        )

    def _save_imported_dataframe(self, dataframe: pd.DataFrame, enterprise: Enterprise) -> tuple[set[str], list[date]]:
        imported_dates: list[date] = []
        records_to_save: list[FinancialRecord] = []

        for row_index, row in dataframe.iterrows():
            try:
                numeric_values = {
                    "revenue": float(row["revenue"]),
                    "cost": float(row["cost"]),
                    "fixed_expenses": float(row["fixed_expenses"]),
                    "variable_expenses": float(row["variable_expenses"]),
                    "tax": float(row["tax"]),
                }
                self._validate_financial_values(numeric_values)
                period_date = row["period_date"]
                record = FinancialRecord(
                    id=None,
                    enterprise_id=enterprise.id,
                    period_date=period_date,
                    revenue=numeric_values["revenue"],
                    cost=numeric_values["cost"],
                    fixed_expenses=numeric_values["fixed_expenses"],
                    variable_expenses=numeric_values["variable_expenses"],
                    tax=numeric_values["tax"],
                )
                records_to_save.append(record)
                imported_dates.append(period_date)
            except Exception as exc:
                raise ValueError(f"Ошибка в строке {row_index + 2}: {exc}") from exc

        self.repository.upsert_records(records_to_save)
        return {enterprise.name}, imported_dates

    def open_record_dialog(
        self,
        enterprise: Enterprise,
        record: FinancialRecord | None = None,
    ) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Редактировать данные" if record else "Добавить данные")
        self._style_dialog(dialog, size="460x360")
        dialog.columnconfigure(0, weight=0)
        dialog.columnconfigure(1, weight=1)

        entries: dict[str, ttk.Entry] = {}
        defaults = self._record_defaults(enterprise, record)

        for row_index, (label, field_name) in enumerate(RECORD_FIELDS):
            ttk.Label(dialog, text=label, style="Dialog.TLabel").grid(row=row_index, column=0, padx=12, pady=6, sticky=tk.W)
            entry = ttk.Entry(dialog, width=22)
            entry.grid(row=row_index, column=1, padx=12, pady=6, sticky=tk.W)
            entry.insert(0, defaults[field_name])
            entries[field_name] = entry

        def save() -> None:
            try:
                payload = self._build_record_from_entries(enterprise.id, entries, record)
                existing = self.repository.get_record_by_date(enterprise.id, payload.period_date)
                if existing and existing.id != payload.id:
                    raise ValueError(
                        f"Данные за {payload.period_date.strftime('%Y-%m-%d')} уже существуют"
                    )
                if payload.id is None:
                    self.repository.add_record(payload)
                else:
                    self.repository.update_record(payload)
            except ValueError as exc:
                self._show_styled_error_dialog("Ошибка", str(exc))
                return
            except psycopg.Error as exc:
                self._show_styled_error_dialog("Ошибка PostgreSQL", str(exc))
                return

            self.ensure_period_includes(payload.period_date)
            self._show_styled_info_dialog("Успех", "Данные успешно сохранены")
            self._close_modal_dialog(dialog)
            self.refresh_current_view()

        button_frame = ttk.Frame(dialog, style="Dialog.TFrame")
        button_frame.grid(row=len(RECORD_FIELDS), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Сохранить", style="Dialog.TButton", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Отмена",
            style="Dialog.TButton",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(side=tk.LEFT, padx=5)
        self._show_modal_dialog(dialog)
        entries["period_date"].focus_set()

    def _record_defaults(
        self,
        enterprise: Enterprise,
        record: FinancialRecord | None,
    ) -> dict[str, str]:
        if record:
            return {
                "period_date": record.period_date.strftime("%Y-%m-%d"),
                "revenue": str(record.revenue),
                "cost": str(record.cost),
                "fixed_expenses": str(record.fixed_expenses),
                "variable_expenses": str(record.variable_expenses),
                "tax": str(record.tax),
            }

        latest_date = self.repository.get_latest_record_date(enterprise.id)
        next_date = latest_date + relativedelta(months=1) if latest_date else date(2025, 1, 1)
        if next_date > date.today():
            next_date = date.today()
        return {
            "period_date": next_date.strftime("%Y-%m-%d"),
            "revenue": "",
            "cost": "",
            "fixed_expenses": "",
            "variable_expenses": "",
            "tax": "",
        }

    def _build_record_from_entries(
        self,
        enterprise_id: int,
        entries: dict[str, ttk.Entry],
        record: FinancialRecord | None,
    ) -> FinancialRecord:
        parsed: dict[str, float | date] = {}
        for field_name, entry in entries.items():
            raw_value = entry.get().strip()
            if not raw_value:
                raise ValueError("Все поля должны быть заполнены")
            if field_name == "period_date":
                try:
                    parsed[field_name] = datetime.strptime(raw_value, "%Y-%m-%d").date()
                except ValueError as exc:
                    raise ValueError("Дата должна быть в формате ГГГГ-ММ-ДД") from exc
            else:
                try:
                    parsed[field_name] = float(raw_value)
                except ValueError as exc:
                    display_name = FIELD_DISPLAY_NAMES.get(field_name, field_name)
                    raise ValueError(f"Поле «{display_name}» должно быть числом") from exc

        numeric_values = {
            "revenue": float(parsed["revenue"]),
            "cost": float(parsed["cost"]),
            "fixed_expenses": float(parsed["fixed_expenses"]),
            "variable_expenses": float(parsed["variable_expenses"]),
            "tax": float(parsed["tax"]),
        }
        self._validate_financial_values(numeric_values)
        self._ensure_not_future_date(parsed["period_date"], "\u0414\u0430\u0442\u0430 \u0437\u0430\u043f\u0438\u0441\u0438")

        return FinancialRecord(
            id=record.id if record else None,
            enterprise_id=enterprise_id,
            period_date=parsed["period_date"],
            revenue=numeric_values["revenue"],
            cost=numeric_values["cost"],
            fixed_expenses=numeric_values["fixed_expenses"],
            variable_expenses=numeric_values["variable_expenses"],
            tax=numeric_values["tax"],
        )

    def _validate_financial_values(self, values: dict[str, float]) -> None:
        for field_name in ("cost", "fixed_expenses", "variable_expenses", "tax"):
            if values[field_name] < 0:
                raise ValueError(f"{FIELD_DISPLAY_NAMES[field_name]} не может быть отрицательной")

    def on_tree_select(self, _event) -> None:
        if self.suppress_tree_event:
            return
        selection = self.tree.selection()
        self.selected_record_id = int(selection[0]) if selection else None
        if self.current_records:
            self.plot_graphs(preserve_view=True)

    def on_enterprise_change(self, _event) -> None:
        self.clear_statistics()
        enterprise = self.get_selected_enterprise()
        if enterprise:
            self.set_period_to_full_range(enterprise)
        self.show_data()

    def _setup_combobox_behavior(self, combobox: ttk.Combobox) -> None:
        combobox.configure(exportselection=False)
        combobox.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.after_idle(self.focus_set),
            add="+",
        )
        combobox.bind(
            "<FocusOut>",
            lambda _event, cb=combobox: self.after_idle(lambda: self._clear_combobox_selection(cb)),
            add="+",
        )

    def _clear_combobox_selection(self, combobox: ttk.Combobox) -> None:
        try:
            combobox.selection_clear()
            combobox.icursor(tk.END)
        except tk.TclError:
            return

    def on_graph_tick_mode_change(self, _event) -> None:
        if self.current_records:
            self.plot_graphs()

    def on_graph_scroll(self, event) -> None:
        if not self.current_records or not self.metrics_axes:
            return

        ax1, ax2 = self.metrics_axes
        if event.inaxes not in (ax1, ax2) or event.xdata is None:
            return

        x_min, x_max = ax1.get_xlim()
        if x_max <= x_min:
            return

        scale_factor = 0.8 if event.button == "up" else 1.25
        cursor_x = event.xdata
        new_x_min = cursor_x - (cursor_x - x_min) * scale_factor
        new_x_max = cursor_x + (x_max - cursor_x) * scale_factor

        data_min = 0.5
        data_max = len(self.current_records) + 0.5
        min_window = 2.0

        if new_x_max - new_x_min < min_window:
            midpoint = cursor_x
            new_x_min = midpoint - min_window / 2
            new_x_max = midpoint + min_window / 2

        if new_x_min < data_min:
            shift = data_min - new_x_min
            new_x_min += shift
            new_x_max += shift
        if new_x_max > data_max:
            shift = new_x_max - data_max
            new_x_min -= shift
            new_x_max -= shift

        new_x_min = max(new_x_min, data_min)
        new_x_max = min(new_x_max, data_max)
        if new_x_max - new_x_min < min_window:
            new_x_min = data_min
            new_x_max = min(data_max, data_min + min_window)

        ax1.set_xlim(new_x_min, new_x_max)
        ax2.set_xlim(new_x_min, new_x_max)
        self._update_graph_x_ticks()
        self.canvas_metrics.draw_idle()

    def on_graph_press(self, event) -> None:
        if not self.current_records or not self.metrics_axes or event.button != 1:
            return

        ax1, _ax2 = self.metrics_axes
        if event.inaxes not in self.metrics_axes or event.x is None:
            return

        x_min, x_max = ax1.get_xlim()
        self.drag_state = {
            "start_x": event.x,
            "x_min": x_min,
            "x_max": x_max,
            "start_y": event.y,
        }
        if self.graph_widget:
            self.graph_widget.configure(cursor="fleur")

    def on_graph_drag(self, event) -> None:
        if not self.drag_state or not self.metrics_axes or event.x is None:
            return

        ax1, ax2 = self.metrics_axes
        axis_bbox = ax1.get_window_extent()
        axis_width = axis_bbox.width
        if axis_width <= 0:
            return

        data_width = self.drag_state["x_max"] - self.drag_state["x_min"]
        delta_pixels = self.drag_state["start_x"] - event.x
        delta_data = (delta_pixels / axis_width) * data_width
        new_x_min = self.drag_state["x_min"] + delta_data
        new_x_max = self.drag_state["x_max"] + delta_data

        data_min = 0.5
        data_max = len(self.current_records) + 0.5
        window_width = data_width

        if new_x_min < data_min:
            new_x_min = data_min
            new_x_max = data_min + window_width
        if new_x_max > data_max:
            new_x_max = data_max
            new_x_min = data_max - window_width

        ax1.set_xlim(new_x_min, new_x_max)
        ax2.set_xlim(new_x_min, new_x_max)
        self._update_graph_x_ticks()
        self.canvas_metrics.draw_idle()

    def on_graph_release(self, event) -> None:
        if self.drag_state and event and event.x is not None and event.y is not None:
            delta_x = abs(event.x - self.drag_state["start_x"])
            delta_y = abs(event.y - self.drag_state["start_y"])
            if delta_x <= 4 and delta_y <= 4 and event.xdata is not None:
                self._select_nearest_record(event.xdata)
        self.drag_state = None
        if self.graph_widget:
            self.graph_widget.configure(cursor="hand2")

    def on_graph_leave(self, _event) -> None:
        if self.graph_widget and not self.drag_state:
            self.graph_widget.configure(cursor="arrow")

    def _select_nearest_record(self, x_value: float) -> None:
        if not self.current_records:
            return
        index = int(round(x_value)) - 1
        self._select_record_by_index(index)

    def _report_metric_values_from_analysis(self) -> dict[str, float]:
        if not self.analysis_result:
            return {}
        return {
            "target_ros": float(self.analysis_result["target_ros"]),
            "alpha": float(self.analysis_result["alpha"]),
            "avg_ros": float(self.analysis_result["avg_ros"]),
            "std_ros": float(self.analysis_result["std_ros"]),
            "min_ros": float(self.analysis_result["min_ros"]),
            "max_ros": float(self.analysis_result["max_ros"]),
            "avg_profit": float(self.analysis_result["avg_profit"]),
            "t_stat": float(self.analysis_result["t_stat"]),
            "p_value": float(self.analysis_result["p_value"]),
        }

    def _save_current_report_to_database(self) -> FinancialReport | None:
        if not self.analysis_result or not self.current_enterprise or not self.current_records:
            return None
        target_ros = float(self.analysis_result.get("target_ros", 0.0))
        alpha = float(self.analysis_result.get("alpha", 0.05))
        report = FinancialReport(
            id=None,
            enterprise_id=self.current_enterprise.id,
            name=f"Аналитический отчёт (ROS {target_ros:.1f}%; α {alpha:.2f})",
            date_created=datetime.now().date(),
            period_start=self.current_records[0].period_date,
            period_end=self.current_records[-1].period_date,
        )
        return self.repository.save_financial_report(report, self._report_metric_values_from_analysis())

    def save_report(self) -> None:
        if not self.analysis_result:
            self._show_styled_warning_dialog("Ошибка", "Сначала проведите проверку гипотезы")
            return

        if self.current_report is None:
            try:
                self.current_report = self._save_current_report_to_database()
            except psycopg.Error as exc:
                self._show_styled_warning_dialog(
                    "Автосохранение отчёта",
                    "Не удалось сохранить отчёт в базу данных перед экспортом.\n\n"
                    f"Текст ошибки: {exc}",
                )
            else:
                if self.current_report:
                    self.analysis_result["report_id"] = self.current_report.id

        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
                filetypes=[("HTML файл", "*.html"), ("Текстовый файл", "*.txt")],
            title="Сохранить аналитический отчёт",
        )
        if not file_path:
            return

        report_html = self._build_analysis_report_html()
        with open(file_path, "w", encoding="utf-8") as report_file:
            report_file.write(report_html)

        self._show_styled_info_dialog("Успех", f"Отчёт сохранён:\n{file_path}")

    def _load_saved_report(self, report_id: int) -> None:
        report = self.repository.get_financial_report(report_id)
        if report is None:
            self._show_styled_error_dialog("Ошибка", "Отчёт не найден")
            return

        metric_values = self.repository.get_financial_report_metric_values(report_id)
        analysis_result = self._analysis_result_from_report(report, metric_values)
        enterprise = next((item for item in self.repository.list_enterprises() if item.id == report.enterprise_id), None)
        if enterprise is None:
            self._show_styled_error_dialog("Ошибка", "Предприятие для отчёта не найдено")
            return

        self.enterprise_var.set(enterprise.name)
        self.period_mode_var.set("Период")
        self._update_period_mode_ui()
        self.period_start_var.set(report.period_start.strftime("%Y-%m-%d"))
        self.period_end_var.set(report.period_end.strftime("%Y-%m-%d"))
        self.show_data()
        self.current_report = report
        self.analysis_result = analysis_result
        self.analysis_result["report_id"] = report.id
        self.render_statistics(self._statistics_lines_from_analysis_result(analysis_result))

from __future__ import annotations

import re
import tkinter as tk
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import psycopg
from dateutil.relativedelta import relativedelta
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from .analysis import summarize_ros, t_test_one_sample
from .models import Enterprise, FinancialRecord
from .repository import PostgresRepository


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


class RentabilityAnalysisApp(tk.Tk):
    def __init__(self, repository: PostgresRepository | None = None):
        super().__init__()
        self.title("ОРМП: Оценка рентабельности малого предприятия")
        self.geometry("1400x800")

        self.repository = repository or PostgresRepository()
        try:
            self.repository.initialize()
        except psycopg.Error as exc:
            messagebox.showerror(
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
        self.selected_record_id: int | None = None
        self.modal_overlay: tk.Toplevel | None = None

        self.enterprise_var = tk.StringVar()
        self.period_mode_var = tk.StringVar(value="Все данные")
        self.period_start_var = tk.StringVar()
        self.period_end_var = tk.StringVar()
        self.target_ros_var = tk.StringVar(value="10.0")
        self.alpha_var = tk.StringVar(value="0.05")
        self.graph_tick_mode_var = tk.StringVar(value="Авто")

        self.fig_metrics: Figure | None = None
        self.canvas_metrics: FigureCanvasTkAgg | None = None
        self.metrics_axes: tuple | None = None
        self.drag_state: dict[str, float] | None = None
        self.graph_widget: tk.Widget | None = None
        self.suppress_tree_event = False

        self.create_widgets()
        self.update_enterprise_list()

    def create_widgets(self) -> None:
        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_panel = ttk.Frame(main_paned)
        right_panel = ttk.Frame(main_paned)
        main_paned.add(left_panel, weight=1)
        main_paned.add(right_panel, weight=2)

        self._create_table_panel(left_panel)
        self._create_settings_panel(right_panel)
        self._create_graph_panel(right_panel)

    def _create_table_panel(self, parent: ttk.Frame) -> None:
        table_frame = ttk.LabelFrame(parent, text="Финансовые данные за период")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(table_frame, columns=TABLE_COLUMNS, show="headings", height=12)
        for column, width in zip(TABLE_COLUMNS, [90, 100, 100, 100, 100, 80, 100, 70]):
            self.tree.heading(column, text=column)
            anchor = tk.W if column == "Дата" else tk.CENTER
            self.tree.column(column, width=width, anchor=anchor)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        y_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        data_actions_frame = ttk.Frame(parent)
        data_actions_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(data_actions_frame, text="Добавить данные", command=self.add_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_actions_frame, text="Редактировать", command=self.edit_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_actions_frame, text="Импорт", command=self.import_data).pack(side=tk.LEFT, padx=2)

        results_frame = ttk.LabelFrame(parent, text="Статистические показатели")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        results_actions = ttk.Frame(results_frame)
        results_actions.pack(fill=tk.X, padx=5, pady=(5, 0))
        ttk.Button(
            results_actions,
            text="Рассчитать показатели",
            command=self.test_hypothesis,
        ).pack(side=tk.LEFT, padx=2)

        self.results_text = tk.Text(
            results_frame,
            height=11,
            font=("Segoe UI", 11),
            state=tk.DISABLED,
            wrap=tk.WORD,
            padx=8,
            pady=8,
        )
        results_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scrollbar.set)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_settings_panel(self, parent: ttk.Frame) -> None:
        settings_frame = ttk.LabelFrame(parent, text="Параметры анализа")
        settings_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

        enterprise_frame = ttk.Frame(settings_frame)
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

        row1 = ttk.Frame(settings_frame)
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

        self.period_range_frame = ttk.Frame(row1)
        self.period_range_frame.pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(self.period_range_frame, text="с").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(self.period_range_frame, textvariable=self.period_start_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.period_range_frame, text="по").pack(side=tk.LEFT, padx=5)
        ttk.Entry(self.period_range_frame, textvariable=self.period_end_var, width=12).pack(side=tk.LEFT, padx=5)
        self._update_period_mode_ui()

        row2 = ttk.Frame(settings_frame)
        row2.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row2, text="Целевой ROS (%):").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.target_ros_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Уровень α:").pack(side=tk.LEFT, padx=(20, 5))
        ttk.Entry(row2, textvariable=self.alpha_var, width=6).pack(side=tk.LEFT, padx=5)

        button_frame = ttk.Frame(settings_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        ttk.Button(button_frame, text="Вывести данные", command=self.show_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Сохранить отчёт", command=self.save_report).pack(side=tk.LEFT, padx=2)

    def _create_graph_panel(self, parent: ttk.Frame) -> None:
        graph_frame = ttk.LabelFrame(parent, text="")
        graph_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        graph_controls = ttk.Frame(graph_frame)
        graph_controls.pack(fill=tk.X, padx=8, pady=(6, 0))
        ttk.Label(graph_controls, text="Подписи месяцев:").pack(side=tk.LEFT)
        tick_mode_combo = ttk.Combobox(
            graph_controls,
            textvariable=self.graph_tick_mode_var,
            state="readonly",
            width=20,
            values=("Авто", "Каждый месяц", "Через 2 месяца"),
        )
        tick_mode_combo.pack(side=tk.LEFT, padx=6)
        tick_mode_combo.bind("<<ComboboxSelected>>", self.on_graph_tick_mode_change)
        ttk.Label(
            graph_controls,
            text="Авто / Каждый месяц / Через 2 месяца",
        ).pack(side=tk.LEFT, padx=6)

        self.fig_metrics = Figure(figsize=(8, 6), dpi=100)
        self.canvas_metrics = FigureCanvasTkAgg(self.fig_metrics, master=graph_frame)
        self.graph_widget = self.canvas_metrics.get_tk_widget()
        self.graph_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=(4, 5))
        self.graph_widget.configure(cursor="hand2")
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

    def get_selected_enterprise(self) -> Enterprise | None:
        enterprise = self.enterprise_by_name.get(self.enterprise_var.get())
        if not enterprise:
            messagebox.showwarning("Ошибка", "Выберите предприятие")
            return None
        return enterprise

    def clear_results(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.selected_record_id = None
        self.current_records = []
        self.analysis_result = None

        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        self.results_text.config(state=tk.DISABLED)

        self.clear_graphs_only()

    def clear_statistics(self) -> None:
        self.analysis_result = None
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
        overlay.attributes("-alpha", 0.25)
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
        self._center_dialog(dialog)
        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: self._close_modal_dialog(dialog))
        dialog.bind("<Escape>", lambda _event: self._close_modal_dialog(dialog))
        dialog.lift()
        dialog.focus_force()

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
            except ValueError as exc:
                messagebox.showerror(
                    "Ошибка даты",
                    "Некорректный формат даты.\n"
                    "Используйте ДД.ММ.ГГГГ или ГГГГ-ММ-ДД.\n\n"
                    f"Текст ошибки: {exc}",
                )
                self.clear_results()
                return

        self.current_enterprise = enterprise
        self.current_records = self.repository.get_records(enterprise.id, start_date, end_date)
        self.tree.delete(*self.tree.get_children())

        for record in self.current_records:
            self.tree.insert(
                "",
                tk.END,
                iid=str(record.id),
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
            self.period_range_frame.pack(side=tk.LEFT, padx=(10, 0))
        else:
            self.period_range_frame.pack_forget()

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
            ax1.axvline(selected_x, color="#6c757d", linestyle=":", linewidth=1.2, alpha=0.7, zorder=0)

        ax1.set_xlabel("Период", fontsize=9)
        ax1.set_ylabel("Чистая прибыль, ₽", fontsize=9)
        ax2.set_ylabel("ROS, %", fontsize=9)
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
            fontsize=8,
            frameon=True,
            borderaxespad=0.3,
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
            messagebox.showwarning("Ошибка", "Недостаточно данных для проверки гипотезы")
            return

        try:
            target_ros = float(self.target_ros_var.get())
            alpha = float(self.alpha_var.get())
            if not 0 < alpha < 1:
                raise ValueError("α должен быть в диапазоне (0, 1)")
            if target_ros < 0:
                raise ValueError("Целевой ROS не может быть отрицательным")
        except ValueError as exc:
            messagebox.showerror("Ошибка ввода", str(exc))
            return

        ros_values = [record.ros for record in self.current_records]
        try:
            t_stat, p_value = t_test_one_sample(ros_values, target_ros)
            avg_ros, std_ros = summarize_ros(ros_values)
        except ValueError as exc:
            messagebox.showerror("Ошибка вычислений", str(exc))
            return

        verdict = "Не отклоняется" if p_value >= alpha else "Отклоняется"
        recommendation = (
            "Рекомендуется к инвестированию"
            if p_value >= alpha
            else "Не рекомендуется к инвестированию"
        )
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
            "t_stat": t_stat,
            "p_value": p_value,
            "verdict": verdict,
            "recommendation": recommendation,
            "target_ros": target_ros,
            "alpha": alpha,
            "enterprise": self.enterprise_var.get(),
            "date_created": datetime.now().strftime("%d.%m.%Y"),
            "period_start": self.period_start_var.get(),
            "period_end": self.period_end_var.get(),
        }

    def save_report(self) -> None:
        if not self.analysis_result:
            messagebox.showwarning("Ошибка", "Сначала проведите проверку гипотезы")
            return

        data_frame = pd.DataFrame(
            [
                {
                    "Дата": record.period_date.strftime("%d.%m.%Y"),
                    "Выручка, ₽": record.revenue,
                    "Себестоимость, ₽": record.cost,
                    "Пост. издержки, ₽": record.fixed_expenses,
                    "Перем. издержки, ₽": record.variable_expenses,
                    "Налог, ₽": record.tax,
                    "Чистая прибыль, ₽": record.net_profit,
                    "ROS, %": record.ros,
                }
                for record in self.current_records
            ]
        )
        summary_frame = pd.DataFrame(
            [
                {
                    "Название предприятия": self.analysis_result["enterprise"],
                    "Дата анализа": self.analysis_result["date_created"],
                    "Период анализа: с": self.analysis_result["period_start"],
                    "Период анализа: по": self.analysis_result["period_end"],
                    "Целевой уровень ROS": f'{self.analysis_result["target_ros"]}%',
                    "Уровень значимости α": self.analysis_result["alpha"],
                    "Средняя ROS": f'{self.analysis_result["avg_ros"]:.1f}%',
                    "Стандартное отклонение": f'{self.analysis_result["std_ros"]:.1f}%',
                    "t-статистика": round(self.analysis_result["t_stat"], 2),
                    "p-уровень": round(self.analysis_result["p_value"], 3),
                    "Вердикт гипотезы": self.analysis_result["verdict"],
                    "Рекомендация": self.analysis_result["recommendation"],
                }
            ]
        )

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить аналитический отчёт",
        )
        if not file_path:
            return

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            data_frame.to_excel(writer, sheet_name="Данные", index=False)
            summary_frame.T.to_excel(writer, sheet_name="Итоги", header=False)

        messagebox.showinfo("Успех", f"Отчёт сохранён:\n{file_path}")

    def add_enterprise(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Добавить предприятие")
        dialog.geometry("320x120")

        ttk.Label(dialog, text="Название предприятия:").pack(pady=8)
        name_entry = ttk.Entry(dialog, width=34)
        name_entry.pack(pady=5)

        def save() -> None:
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("Ошибка", "Введите название предприятия")
                return
            if name in self.enterprise_by_name:
                messagebox.showerror("Ошибка", "Предприятие с таким названием уже существует")
                return
            try:
                enterprise = self.repository.add_enterprise(name)
            except psycopg.errors.UniqueViolation:
                messagebox.showerror("Ошибка", "Предприятие с таким названием уже существует")
                return
            self.enterprise_by_name[enterprise.name] = enterprise
            self.update_enterprise_list()
            self.enterprise_var.set(enterprise.name)
            self.clear_statistics()
            self._close_modal_dialog(dialog)
            self.refresh_current_view(rerun_analysis=False)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=12)
        ttk.Button(button_frame, text="Сохранить", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Отмена",
            command=lambda: self._close_modal_dialog(dialog),
        ).pack(side=tk.LEFT, padx=5)
        self._show_modal_dialog(dialog, use_overlay=True)
        name_entry.focus_set()

    def add_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if enterprise:
            self.open_record_dialog(enterprise)

    def delete_enterprise(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise:
            return

        confirmed = messagebox.askyesno(
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
                confirmed = messagebox.askyesno(
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
            messagebox.showerror("Ошибка импорта", str(exc))
            return
        except psycopg.Error as exc:
            messagebox.showerror("Ошибка PostgreSQL", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Ошибка импорта", f"Не удалось импортировать файл:\n{exc}")
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
        messagebox.showinfo(
            "Импорт завершён",
            f"Импортировано строк: {len(imported_dates)}\n"
            f"Предприятие: {enterprise.name}",
        )

    def edit_data(self) -> None:
        enterprise = self.get_selected_enterprise()
        if not enterprise or self.selected_record_id is None:
            messagebox.showwarning("Ошибка", "Выберите строку для редактирования")
            return

        record = next((item for item in self.current_records if item.id == self.selected_record_id), None)
        if not record:
            messagebox.showerror("Ошибка", "Данные не найдены")
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
        dataframe["period_date"] = period_dates

        return dataframe

    @staticmethod
    def _normalize_import_column_name(column: object) -> str:
        normalized_key = str(column).strip().lower().replace("ё", "е")
        normalized_key = normalized_key.replace("₽", "").replace("%", "")
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
                period_date = row["period_date"]
                record = FinancialRecord(
                    id=None,
                    enterprise_id=enterprise.id,
                    period_date=period_date,
                    revenue=float(row["revenue"]),
                    cost=float(row["cost"]),
                    fixed_expenses=float(row["fixed_expenses"]),
                    variable_expenses=float(row["variable_expenses"]),
                    tax=float(row["tax"]),
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
        dialog.geometry("420x320")

        entries: dict[str, ttk.Entry] = {}
        defaults = self._record_defaults(enterprise, record)

        for row_index, (label, field_name) in enumerate(RECORD_FIELDS):
            ttk.Label(dialog, text=label).grid(row=row_index, column=0, padx=10, pady=5, sticky=tk.W)
            entry = ttk.Entry(dialog, width=22)
            entry.grid(row=row_index, column=1, padx=10, pady=5, sticky=tk.W)
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
                messagebox.showerror("Ошибка", str(exc))
                return
            except psycopg.Error as exc:
                messagebox.showerror("Ошибка PostgreSQL", str(exc))
                return

            self.ensure_period_includes(payload.period_date)
            messagebox.showinfo("Успех", "Данные успешно сохранены")
            self._close_modal_dialog(dialog)
            self.refresh_current_view()

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=len(RECORD_FIELDS), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Сохранить", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Отмена",
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
                    raise ValueError(f"Поле {field_name} должно быть числом") from exc

        return FinancialRecord(
            id=record.id if record else None,
            enterprise_id=enterprise_id,
            period_date=parsed["period_date"],
            revenue=float(parsed["revenue"]),
            cost=float(parsed["cost"]),
            fixed_expenses=float(parsed["fixed_expenses"]),
            variable_expenses=float(parsed["variable_expenses"]),
            tax=float(parsed["tax"]),
        )

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

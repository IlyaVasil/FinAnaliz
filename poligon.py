import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import re


class FinancialAnalyzer:
    """
    Основной класс приложения для финансового анализа предприятия.
    Управляет интерфейсом, вводом данных, расчётами и визуализацией.
    """

    def __init__(self, main_window):
        """
        Инициализация основного окна приложения: заголовок, иконка, размер, стили, интерфейс.
        """
        self.window = main_window
        self.window.title("Финансовый анализ предприятия")

        # Попытка установить иконку из файла icon.ico
        icon_path = "icon.ico"
        if os.path.isfile(icon_path):
            self.window.iconbitmap(icon_path)
        else:
            print("Предупреждение: файл иконки 'icon.ico' не найден.")

        # Установка размера окна ~90% от экрана и центрирование
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.9)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Настройка внешнего вида виджетов
        self.setup_styles()

        # Словарь для хранения полей ввода
        self.input_fields = {}
        # Результаты анализа (заполняется после расчёта)
        self.analysis_results = None
        # Текущая активная вкладка результатов
        self.current_view = None

        # Создание меню и основного интерфейса
        self.build_menu()
        self.build_interface()

    def setup_styles(self):
        """
        Настраивает внешний вид элементов интерфейса с использованием ttk.Style.
        Включает шрифты, отступы и специальный стиль для маленьких кнопок.
        """
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 11), padding=6)
        style.configure('TEntry', padding=4, font=('Segoe UI', 11))
        style.configure('TLabelframe.Label', font=('Segoe UI', 11, 'bold'))
        style.configure('TNotebook.Tab', font=('Segoe UI', 11, 'bold'), padding=[12, 6])
        style.configure('Small.TButton', font=('Segoe UI', 9), padding=4)

    def build_menu(self):
        """
        Создаёт главное меню приложения с пунктами: Файл, Инструменты, Справка.
        Каждый пункт содержит подкоманды для загрузки, экспорта, анализа и помощи.
        """
        menu_bar = tk.Menu(self.window)
        self.window.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Загрузить из Excel", command=self.load_excel)
        file_menu.add_command(label="Скачать шаблон", command=self.make_template)
        file_menu.add_separator()
        file_menu.add_command(label="Сохранить в Excel", command=self.export_excel)
        file_menu.add_command(label="Сохранить в PDF", command=self.export_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.window.quit)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        tools_menu = tk.Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Провести расчёт", command=self.run_analysis)
        menu_bar.add_cascade(label="Инструменты", menu=tools_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="О программе", command=self.show_info)
        help_menu.add_separator()
        help_menu.add_command(label="Назначение и возможности программы", command=self.show_purpose)
        help_menu.add_command(label="Начало работы", command=self.show_getting_started)
        help_menu.add_command(label="Аналитические возможности программы", command=self.show_analytics)
        help_menu.add_command(label="Импорт и экспорт данных", command=self.show_import_export)
        help_menu.add_command(label="Настройки", command=self.show_settings_help)
        help_menu.add_command(label="Использование программы", command=self.show_usage)
        menu_bar.add_cascade(label="Справка", menu=help_menu)

    def show_info(self):
        """Показывает диалоговое окно с информацией о программе и авторе."""
        messagebox.showinfo("О программе",
                            "Финансовый анализ предприятия\n"
                            "Версия 1.01\n"
                            "Автор: Василенко Илья Константинович\n"
                            "Разработано в 2025 году")

    def show_purpose(self):
        """Отображает назначение и основные возможности программы."""
        text = (
            "НАЗНАЧЕНИЕ И ВОЗМОЖНОСТИ ПРОГРАММЫ\n\n"
            "Программа предназначена для комплексного финансового анализа предприятия "
            "на основе данных бухгалтерской отчётности (Форма №1 и Форма №2).\n\n"
            "Возможности:\n"
            "• Расчёт ключевых финансовых коэффициентов\n"
            "• Анализ ликвидности, рентабельности, финансовой устойчивости и деловой активности\n"
            "• Визуализация результатов в виде графиков\n"
            "• Импорт данных из Excel и экспорт результатов в Excel/PDF"
        )
        messagebox.showinfo("Назначение и возможности", text)

    def show_getting_started(self):
        """Показывает пошаговую инструкцию по началу работы."""
        text = (
            "НАЧАЛО РАБОТЫ\n\n"
            "1. Заполните поля «Название предприятия» и «ИНН» (опционально).\n"
            "2. Введите данные из бухгалтерской отчётности в соответствующие поля.\n"
            "   — Баланс (Форма №1)\n"
            "   — Отчёт о финансовых результатах (Форма №2)\n"
            "3. Нажмите кнопку «Провести финансовый анализ».\n"
            "4. Переключайтесь между вкладками в левой панели для просмотра результатов."
        )
        messagebox.showinfo("Начало работы", text)

    def show_analytics(self):
        """Описывает аналитические возможности: какие коэффициенты рассчитываются."""
        text = (
            "АНАЛИТИЧЕСКИЕ ВОЗМОЖНОСТИ ПРОГРАММЫ\n\n"
            "Программа рассчитывает более 20 финансовых показателей, сгруппированных по разделам:\n\n"
            "ЛИКВИДНОСТЬ:\n"
            "  • Текущая, срочная и абсолютная ликвидность\n"
            "  • Коэффициент обеспеченности собственными оборотными средствами\n\n"
            "РЕТАБЕЛЬНОСТЬ:\n"
            "  • Рентабельность продаж, активов, собственного капитала (ROE), затрат\n"
            "  • EBITDA margin, ROI и др.\n\n"
            "ФИНАНСОВАЯ УСТОЙЧИВОСТЬ:\n"
            "  • Коэффициент автономии, финансовый леверидж, покрытие процентов\n"
            "  • Debt/EBITDA, манёвренность СК\n\n"
            "ДЕЛОВАЯ АКТИВНОСТЬ:\n"
            "  • Оборачиваемость запасов, дебиторской и кредиторской задолженности\n"
            "  • Операционный цикл и др."
        )
        messagebox.showinfo("Аналитические возможности", text)

    def show_import_export(self):
        """Объясняет, как импортировать и экспортировать данные."""
        text = (
            "ИМПОРТ И ЭКСПОРТ ДАННЫХ\n\n"
            "ИМПОРТ:\n"
            "• Используйте «Файл → Загрузить из Excel», чтобы импортировать данные.\n"
            "• Файл должен содержать лист «Данные» с двумя столбцами: «Показатель» и «Значение».\n"
            "• Скачайте шаблон через «Файл → Скачать шаблон» для правильного формата.\n\n"
            "ЭКСПОРТ:\n"
            "• После расчёта сохраните результаты в Excel или PDF через меню «Файл».\n"
            "• Экспорт включает все рассчитанные коэффициенты и информацию о компании."
        )
        messagebox.showinfo("Импорт и экспорт данных", text)

    def show_settings_help(self):
        """Описывает доступные настройки (ограниченные в текущей версии)."""
        text = (
            "НАСТРОЙКИ\n\n"
            "В текущей версии программы настройки не вынесены в отдельное окно.\n"
            "Однако вы можете:\n"
            "• Изменить размер окна вручную\n"
            "• Использовать системные настройки масштабирования (DPI)\n"
            "• Разместить файл иконки «icon.ico» в папке программы для отображения значка\n"
            "• Добавить шрифт «DejaVuSans.ttf» для корректного отображения кириллицы в PDF"
        )
        messagebox.showinfo("Настройки", text)

    def show_usage(self):
        """Даёт практические рекомендации по использованию программы."""
        text = (
            "ИСПОЛЬЗОВАНИЕ ПРОГРАММЫ\n\n"
            "1. Ввод данных:\n"
            "   — Все поля, кроме «Чистая прибыль», обязательны для заполнения.\n"
            "   — Допускаются только неотрицательные числа.\n\n"
            "2. Анализ:\n"
            "   — Нажмите «Провести финансовый анализ» для расчёта показателей.\n"
            "   — Программа автоматически проверит корректность данных.\n\n"
            "3. Просмотр результатов:\n"
            "   — Используйте боковую панель для перехода между разделами анализа.\n"
            "   — Каждый раздел содержит числовые результаты и рекомендации.\n\n"
            "4. Сохранение:\n"
            "   — Экспортируйте результаты в Excel или PDF для отчётов."
        )
        messagebox.showinfo("Использование программы", text)

    def build_interface(self):
        """
        Создаёт основной интерфейс: поля ввода для компании, баланса и отчёта о прибылях,
        кнопку анализа и панель результатов с боковой навигацией.
        """
        # Верхняя панель: название и ИНН
        company_frame = ttk.LabelFrame(self.window, text="Информация о компании")
        company_frame.pack(fill='x', padx=15, pady=5)

        ttk.Label(company_frame, text="Название предприятия:", font=('Segoe UI', 11)).grid(row=0, column=0, sticky='w', padx=10, pady=8)
        self.company_name = ttk.Entry(company_frame, width=60)
        self.company_name.grid(row=0, column=1, padx=10, pady=8, sticky='ew')

        ttk.Label(company_frame, text="ИНН:", font=('Segoe UI', 11)).grid(row=1, column=0, sticky='w', padx=10, pady=8)
        self.inn_field = ttk.Entry(company_frame, width=20)
        self.inn_field.grid(row=1, column=1, padx=10, pady=8, sticky='w')

        company_frame.columnconfigure(1, weight=1)

        # Основная область: баланс и прибыли
        data_frame = ttk.Frame(self.window)
        data_frame.pack(fill='both', expand=True, padx=15, pady=5)

        balance_frame = ttk.LabelFrame(data_frame, text="Баланс (Форма №1)")
        profit_frame = ttk.LabelFrame(data_frame, text="Отчёт о прибылях и убытках (Форма №2)")

        balance_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 8), pady=0)
        profit_frame.grid(row=0, column=1, sticky='nsew', padx=(8, 0), pady=0)

        data_frame.columnconfigure(0, weight=1)
        data_frame.columnconfigure(1, weight=1)
        data_frame.rowconfigure(0, weight=1)

        # Поля баланса
        balance_items = [
            ("Оборотные активы (ОА)", "1200"),
            ("Запасы", "1210"),
            ("Дебиторская задолженность (ДЗ)", "1230"),
            ("Денежные средства (ДС)", "1250"),
            ("Внеоборотные активы (ВА)", "1100"),
            ("Собственный капитал (СК)", "1300"),
            ("Долгосрочные обязательства (ДО)", "1400"),
            ("Кредиторская задолженность (КЗ)", "1520"),
            ("Краткосрочные обязательства (КО)", "1500")
        ]

        for i, (name, code) in enumerate(balance_items):
            ttk.Label(balance_frame, text=f"{name} [{code}]:", font=('Segoe UI', 11)).grid(row=i, column=0, sticky='w', padx=10, pady=6)
            entry = ttk.Entry(balance_frame, width=15)
            entry.grid(row=i, column=1, padx=10, pady=6, sticky='e')
            self.input_fields[name] = entry

        # Поля прибылей
        profit_items = [
            ("Выручка", "2110"),
            ("Чистая прибыль", "2400"),
            ("EBITDA", "2100+Аморт."),
            ("Операционная прибыль (EBIT)", "2200"),
            ("Проценты к уплате", "2300")
        ]
        for i, (name, code) in enumerate(profit_items):
            ttk.Label(profit_frame, text=f"{name} [{code}]:", font=('Segoe UI', 11)).grid(row=i, column=0, sticky='w', padx=10, pady=6)
            entry = ttk.Entry(profit_frame, width=15)
            entry.grid(row=i, column=1, padx=10, pady=6, sticky='e')
            self.input_fields[name] = entry

        # Кнопка запуска анализа
        analyze_btn = ttk.Button(self.window, text="Провести финансовый анализ", command=self.run_analysis)
        analyze_btn.pack(pady=12)

        # Основная панель результатов
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))

        sidebar = ttk.Frame(main_frame, width=180)
        sidebar.pack(side='left', fill='y', padx=(0, 10))
        sidebar.pack_propagate(False)

        self.content_area = ttk.Frame(main_frame)
        self.content_area.pack(side='right', fill='both', expand=True)
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Кнопки навигации по разделам анализа
        self.btn_liquidity = ttk.Button(sidebar, text="ЛИКВИДНОСТЬ", command=self.show_liquidity, style='Small.TButton')
        self.btn_profitability = ttk.Button(sidebar, text="РЕТАБЕЛЬНОСТЬ", command=self.show_profitability, style='Small.TButton')
        self.btn_stability = ttk.Button(sidebar, text="ФИН. УСТОЙЧИВОСТЬ", command=self.show_stability, style='Small.TButton')
        self.btn_activity = ttk.Button(sidebar, text="ДЕЛОВАЯ АКТИВНОСТЬ", command=self.show_activity, style='Small.TButton')

        self.btn_liquidity.pack(fill='x', pady=4)
        self.btn_profitability.pack(fill='x', pady=4)
        self.btn_stability.pack(fill='x', pady=4)
        self.btn_activity.pack(fill='x', pady=4)

        # Создание фреймов для каждого раздела результатов
        self.frame_liquidity = self._create_result_frame()
        self.frame_profitability = self._create_result_frame()
        self.frame_stability = self._create_result_frame()
        self.frame_activity = self._create_result_frame()

        self.frame_liquidity.grid(row=0, column=0, sticky='nsew')
        self.frame_profitability.grid(row=0, column=0, sticky='nsew')
        self.frame_stability.grid(row=0, column=0, sticky='nsew')
        self.frame_activity.grid(row=0, column=0, sticky='nsew')

        # По умолчанию показываем ликвидность
        self.frame_profitability.grid_remove()
        self.frame_stability.grid_remove()
        self.frame_activity.grid_remove()
        self.current_view = self.frame_liquidity

    def _create_result_frame(self):
        """
        Создаёт фрейм для отображения результатов: текстовое поле слева и график справа.
        Возвращает настроенный фрейм с атрибутами text, fig, ax, canvas.
        """
        frame = ttk.Frame(self.content_area)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

        text_widget = tk.Text(
            frame, wrap='word', font=('Consolas', 11),
            bg='#f9f9f9', relief='flat', padx=12, pady=12,
            height=8
        )
        text_widget.grid(row=0, column=0, sticky='nsew', padx=(10, 5), pady=10)
        text_widget.config(state='disabled')

        fig = plt.Figure(figsize=(5.5, 3.2), tight_layout=True)
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=0, column=1, sticky='nsew', padx=(5, 10), pady=10)

        frame.text = text_widget
        frame.fig = fig
        frame.ax = ax
        frame.canvas = canvas
        frame.canvas_widget = canvas_widget

        return frame

    def _switch_view(self, frame):
        """
        Переключает отображаемый фрейм результатов (ликвидность, рентабельность и т.д.).
        Также запускает отрисовку графика с небольшой задержкой для корректного размера.
        """
        if self.current_view:
            self.current_view.grid_remove()
        frame.grid(row=0, column=0, sticky='nsew')
        self.current_view = frame
        self.window.after(100, lambda: self._ensure_canvas_draw(frame))

    def _ensure_canvas_draw(self, frame):
        """
        Гарантирует, что график будет нарисован только после того,
        как его виджет получит корректные размеры. Использует рекурсивную проверку.
        """
        if not hasattr(frame, 'canvas') or not self.analysis_results:
            return

        self.window.update_idletasks()
        w = frame.canvas_widget.winfo_width()
        h = frame.canvas_widget.winfo_height()

        if w <= 1 or h <= 1:
            frame.canvas_widget.config(width=400, height=200)
            self.window.after(50, lambda: self._ensure_canvas_draw(frame))
            return

        # Обновление содержимого в зависимости от текущего фрейма
        if frame == self.frame_liquidity:
            self.update_liquidity_view()
        elif frame == self.frame_profitability:
            self.update_profitability_view()
        elif frame == self.frame_stability:
            self.update_stability_view()
        elif frame == self.frame_activity:
            self.update_activity_view()

    def show_liquidity(self):
        """Переключает интерфейс на вкладку «Ликвидность»."""
        self._switch_view(self.frame_liquidity)

    def show_profitability(self):
        """Переключает интерфейс на вкладку «Рентабельность»."""
        self._switch_view(self.frame_profitability)

    def show_stability(self):
        """Переключает интерфейс на вкладку «Финансовая устойчивость»."""
        self._switch_view(self.frame_stability)

    def show_activity(self):
        """Переключает интерфейс на вкладку «Деловая активность»."""
        self._switch_view(self.frame_activity)

    def check_data(self):
        """
        Проверяет корректность введённых данных:
        - формат ИНН (10 или 12 цифр),
        - обязательные поля заполнены,
        - значения неотрицательные и числовые.
        Возвращает словарь с данными или None в случае ошибки.
        """
        inn = self.inn_field.get().strip()
        if inn and not re.fullmatch(r"\d{10}|\d{12}", inn):
            messagebox.showerror("Ошибка", "ИНН должен содержать 10 или 12 цифр.")
            return None

        required_positive = {
            "Оборотные активы (ОА)",
            "Запасы",
            "Дебиторская задолженность (ДЗ)",
            "Денежные средства (ДС)",
            "Внеоборотные активы (ВА)",
            "Собственный капитал (СК)",
            "Кредиторская задолженность (КЗ)",
            "Краткосрочные обязательства (КО)",
            "Выручка",
            "EBITDA",
            "Операционная прибыль (EBIT)",
            "Проценты к уплате"
        }

        collected = {}
        for field_name, entry in self.input_fields.items():
            raw_value = entry.get().strip()
            if not raw_value:
                if field_name in ["Чистая прибыль"]:
                    raw_value = "0"
                else:
                    messagebox.showerror("Ошибка", f"Поле '{field_name}' обязательно для заполнения.")
                    return None

            try:
                num_value = float(raw_value)
                if num_value < 0:
                    if field_name in required_positive:
                        messagebox.showerror("Ошибка", f"Поле '{field_name}' не может быть отрицательным.")
                        return None
                collected[field_name] = num_value
            except ValueError:
                messagebox.showerror("Ошибка", f"Неверный формат числа в поле '{field_name}'.")
                return None

        collected["Название предприятия"] = self.company_name.get().strip()
        collected["ИНН"] = inn
        return collected

    def run_analysis(self):
        """
        Основной метод расчёта финансовых коэффициентов.
        Извлекает данные, проверяет их, вычисляет показатели по четырём блокам
        и сохраняет результаты в self.analysis_results.
        Затем обновляет все вкладки интерфейса.
        """
        data = self.check_data()
        if not data:
            return

        # Баланс
        oa = data["Оборотные активы (ОА)"]
        reserves = data["Запасы"]
        ar = data["Дебиторская задолженность (ДЗ)"]
        ds = data["Денежные средства (ДС)"]
        va = data["Внеоборотные активы (ВА)"]
        sk = data["Собственный капитал (СК)"]
        do = data["Долгосрочные обязательства (ДО)"]
        ap = data["Кредиторская задолженность (КЗ)"]
        ko = data["Краткосрочные обязательства (КО)"]

        # Прибыли
        revenue = data["Выручка"]
        net_profit = data["Чистая прибыль"]
        ebitda = data["EBITDA"]
        ebit = data["Операционная прибыль (EBIT)"]
        interest = data["Проценты к уплате"]

        total_assets = oa + va
        total_debt = do + ko

        # === ЛИКВИДНОСТЬ ===
        current_liquidity = oa / ko if ko != 0 else 0
        quick_ratio = (oa - reserves) / ko if ko != 0 else 0
        cash_liquidity = ds / ko if ko != 0 else 0
        own_working_capital = sk - va
        sos = own_working_capital / oa if oa != 0 else 0
        coverage_ratio = current_liquidity

        # === РЕНТАБЕЛЬНОСТЬ ===
        sales_profitability = net_profit / revenue if revenue != 0 else 0
        assets_profitability = net_profit / total_assets if total_assets != 0 else 0
        roe = net_profit / sk if sk != 0 else 0
        costs = revenue - net_profit
        profitability_costs = net_profit / costs if costs != 0 else 0
        ebitda_margin = ebitda / revenue if revenue != 0 else 0
        profitability_fixed_assets = net_profit / va if va != 0 else 0
        roe_pre_tax = ebit / sk if sk != 0 else 0
        roi = ebit / (sk + do) if (sk + do) != 0 else 0

        # === ФИНАНСОВАЯ УСТОЙЧИВОСТЬ ===
        autonomy = sk / (sk + total_debt) if (sk + total_debt) != 0 else 0
        dependence = total_debt / (sk + total_debt) if (sk + total_debt) != 0 else 0
        financial_leverage = total_debt / sk if sk != 0 else 0
        interest_coverage = ebit / interest if interest != 0 else float('inf') if ebit > 0 else 0
        asset_coverage = total_assets / total_debt if total_debt != 0 else float('inf')
        maneuverability = (sk - va) / sk if sk != 0 else 0
        debt_to_ebitda = total_debt / ebitda if ebitda != 0 else float('inf')
        long_term_stability = sk / (sk + do) if (sk + do) != 0 else 0
        financial_risk = financial_leverage * (1 - autonomy)

        # === ДЕЛОВАЯ АКТИВНОСТЬ ===
        inventory_turnover = revenue / reserves if reserves != 0 else 0
        ar_turnover = revenue / ar if ar != 0 else 0
        ar_days = 365 / ar_turnover if ar_turnover != 0 else 0
        ap_turnover = revenue / ap if ap != 0 else 0
        assets_turnover = revenue / total_assets if total_assets != 0 else 0
        equity_turnover = revenue / sk if sk != 0 else 0
        inventory_days = 365 / inventory_turnover if inventory_turnover != 0 else 0
        operating_cycle = inventory_days + ar_days

        self.analysis_results = {
            # Ликвидность
            "Текущая ликвидность": current_liquidity,
            "Срочная ликвидность": quick_ratio,
            "Абсолютная ликвидность": cash_liquidity,
            "Коэффициент обеспеченности СОС": sos,
            "Общий коэффициент покрытия": coverage_ratio,

            # Рентабельность
            "Рентабельность продаж": sales_profitability,
            "Рентабельность активов": assets_profitability,
            "ROE": roe,
            "Рентабельность затрат": profitability_costs,
            "EBITDA margin": ebitda_margin,
            "Рентабельность ВА": profitability_fixed_assets,
            "ROE до налогообложения": roe_pre_tax,
            "ROI": roi,

            # Финансовая устойчивость
            "Коэффициент автономии": autonomy,
            "Финансовая зависимость": dependence,
            "Финансовый леверидж": financial_leverage,
            "Покрытие процентов": interest_coverage,
            "Обеспеченность обязательств активами": asset_coverage,
            "Манёвренность СК": maneuverability,
            "Debt-to-EBITDA": debt_to_ebitda,
            "Долгосрочная фин. устойчивость": long_term_stability,
            "Финансовый риск": financial_risk,

            # Деловая активность
            "Оборачиваемость запасов": inventory_turnover,
            "Оборачиваемость ДЗ": ar_turnover,
            "Период погашения ДЗ (дни)": ar_days,
            "Оборачиваемость КЗ": ap_turnover,
            "Оборачиваемость активов": assets_turnover,
            "Оборачиваемость СК": equity_turnover,
            "Операционный цикл": operating_cycle
        }

        # Обновление всех вкладок
        self.update_liquidity_view()
        self.update_profitability_view()
        self.update_stability_view()
        self.update_activity_view()

    def update_liquidity_view(self):
        """
        Обновляет текст и график на вкладке «Ликвидность»:
        - выводит значения коэффициентов,
        - добавляет рекомендации при отклонениях,
        - строит столбчатую диаграмму.
        """
        if not self.analysis_results:
            return

        cl = self.analysis_results["Текущая ликвидность"]
        ql = self.analysis_results["Срочная ликвидность"]
        al = self.analysis_results["Абсолютная ликвидность"]
        sos = self.analysis_results["Коэффициент обеспеченности СОС"]

        liquidity_text = "ЛИКВИДНОСТЬ:\n"
        liquidity_text += f"  Текущая ликвидность: {cl:.2f}\n"
        liquidity_text += f"  Срочная ликвидность: {ql:.2f}\n"
        liquidity_text += f"  Абсолютная ликвидность: {al:.2f}\n"
        liquidity_text += f"  Коэффициент обеспеченности СОС: {sos:.2%}\n\n"

        notes = []
        if cl < 1.0:
            notes.append("⚠ Текущая ликвидность < 1.0 — риск неплатёжеспособности.")
        if ql < 0.7:
            notes.append("⚠ Срочная ликвидность < 0.7 — недостаточно быстрых активов.")
        if al < 0.2:
            notes.append("⚠ Абсолютная ликвидность < 0.2 — не хватает денежных средств.")
        if sos < 0.1:
            notes.append("⚠ Низкая обеспеченность собственными оборотными средствами.")

        if notes:
            liquidity_text += "РЕКОМЕНДАЦИИ:\n" + "\n".join(f"  • {n}" for n in notes)

        self.frame_liquidity.text.config(state='normal')
        self.frame_liquidity.text.delete(1.0, tk.END)
        self.frame_liquidity.text.insert(tk.END, liquidity_text)
        self.frame_liquidity.text.config(state='disabled')

        self.frame_liquidity.ax.clear()
        labels = ["Текущая", "Срочная", "Абсолютная"]
        values = [cl, ql, al]
        colors = ['#4e79a7', '#f28e2b', '#e15759']
        self.frame_liquidity.ax.bar(labels, values, color=colors)
        self.frame_liquidity.ax.axhline(y=1.0, color='gray', linestyle='--')
        self.frame_liquidity.ax.axhline(y=0.7, color='orange', linestyle='--')
        self.frame_liquidity.ax.axhline(y=0.2, color='red', linestyle='--')
        self.frame_liquidity.ax.set_ylabel("Значение")
        self.frame_liquidity.ax.set_title("Коэффициенты ликвидности")
        self.frame_liquidity.ax.grid(axis='y', linestyle='--', alpha=0.5)
        self.frame_liquidity.canvas.draw()

    def update_profitability_view(self):
        """
        Обновляет вкладку «Рентабельность»: показатели, рекомендации и график.
        """
        if not self.analysis_results:
            return

        sp = self.analysis_results["Рентабельность продаж"]
        ap = self.analysis_results["Рентабельность активов"]
        roe = self.analysis_results["ROE"]
        pc = self.analysis_results["Рентабельность затрат"]
        ebitda_m = self.analysis_results["EBITDA margin"]
        roa_fixed = self.analysis_results["Рентабельность ВА"]
        roe_pre = self.analysis_results["ROE до налогообложения"]
        roi = self.analysis_results["ROI"]

        profitability_text = "РЕТАБЕЛЬНОСТЬ:\n"
        profitability_text += f"  Рентабельность продаж: {sp:.2%}\n"
        profitability_text += f"  Рентабельность активов: {ap:.2%}\n"
        profitability_text += f"  ROE: {roe:.2%}\n"
        profitability_text += f"  Рентабельность затрат: {pc:.2%}\n"
        profitability_text += f"  EBITDA margin: {ebitda_m:.2%}\n"
        profitability_text += f"  Рентабельность ВА: {roa_fixed:.2%}\n"
        profitability_text += f"  ROE до налогообложения: {roe_pre:.2%}\n"
        profitability_text += f"  ROI: {roi:.2%}\n\n"

        notes = []
        if sp <= 0:
            notes.append("⚠ Отрицательная рентабельность продаж.")
        if roe <= 0:
            notes.append("⚠ ROE ≤ 0 — капитал не приносит доход.")
        if roi <= 0:
            notes.append("⚠ ROI ≤ 0 — инвестиции убыточны.")

        if notes:
            profitability_text += "РЕКОМЕНДАЦИИ:\n" + "\n".join(f"  • {n}" for n in notes)

        self.frame_profitability.text.config(state='normal')
        self.frame_profitability.text.delete(1.0, tk.END)
        self.frame_profitability.text.insert(tk.END, profitability_text)
        self.frame_profitability.text.config(state='disabled')

        self.frame_profitability.ax.clear()
        labels = ["Продаж", "Активов", "ROE", "ROI"]
        values = [sp, ap, roe, roi]
        colors = ['#59a14f', '#4e79a7', '#f28e2b', '#e15759']
        self.frame_profitability.ax.bar(labels, values, color=colors)
        self.frame_profitability.ax.set_ylabel("Доля")
        self.frame_profitability.ax.set_title("Ключевые показатели рентабельности")
        self.frame_profitability.ax.grid(axis='y', linestyle='--', alpha=0.5)
        self.frame_profitability.ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))
        self.frame_profitability.canvas.draw()

    def update_stability_view(self):
        """
        Обновляет вкладку «Финансовая устойчивость» с горизонтальной диаграммой.
        """
        if not self.analysis_results:
            return

        ar = self.analysis_results["Коэффициент автономии"]
        dep = self.analysis_results["Финансовая зависимость"]
        lev = self.analysis_results["Финансовый леверидж"]
        ic = self.analysis_results["Покрытие процентов"]
        ac = self.analysis_results["Обеспеченность обязательств активами"]
        man = self.analysis_results["Манёвренность СК"]
        d_ebitda = self.analysis_results["Debt-to-EBITDA"]
        lt = self.analysis_results["Долгосрочная фин. устойчивость"]
        risk = self.analysis_results["Финансовый риск"]

        stability_text = "ФИНАНСОВАЯ УСТОЙЧИВОСТЬ:\n"
        stability_text += f"  Коэффициент автономии: {ar:.2%}\n"
        stability_text += f"  Финансовая зависимость: {dep:.2%}\n"
        stability_text += f"  Финансовый леверидж: {lev:.2f}\n"
        stability_text += f"  Покрытие процентов: {ic:.2f}\n"
        stability_text += f"  Обеспеченность обязательств активами: {ac:.2f}\n"
        stability_text += f"  Манёвренность СК: {man:.2%}\n"
        stability_text += f"  Debt-to-EBITDA: {d_ebitda:.2f}\n"
        stability_text += f"  Долгосрочная устойчивость: {lt:.2%}\n"
        stability_text += f"  Финансовый риск: {risk:.2f}\n\n"

        notes = []
        if lev > 2.0:
            notes.append("⚠ Высокий финансовый леверидж (>2) — риск перегрузки долгами.")
        if ic < 3.0:
            notes.append("⚠ Покрытие процентов < 3 — риск дефолта по кредитам.")
        if d_ebitda > 4.0:
            notes.append("⚠ Debt/EBITDA > 4 — долговая нагрузка критична.")
        if man < 0.2:
            notes.append("⚠ Низкая манёвренность СК (<20%) — мало свободных средств.")

        if notes:
            stability_text += "РЕКОМЕНДАЦИИ:\n" + "\n".join(f"  • {n}" for n in notes)

        self.frame_stability.text.config(state='normal')
        self.frame_stability.text.delete(1.0, tk.END)
        self.frame_stability.text.insert(tk.END, stability_text)
        self.frame_stability.text.config(state='disabled')

        self.frame_stability.ax.clear()
        labels = [
            "Автономия", "Леверидж", "Покрытие %",
            "Debt/EBITDA", "Манёвренность"
        ]
        values = [ar, lev, ic, d_ebitda, man]
        colors = ['#59a14f', '#e15759', '#4e79a7', '#f28e2b', '#b07aa1']

        y_pos = range(len(labels))
        self.frame_stability.ax.barh(y_pos, values, color=colors, height=0.6)
        self.frame_stability.ax.set_yticks(y_pos)
        self.frame_stability.ax.set_yticklabels(labels, fontsize=10, ha='right')
        self.frame_stability.ax.set_xlabel("Значение")
        self.frame_stability.ax.set_title("Показатели финансовой устойчивости")
        self.frame_stability.ax.grid(axis='x', linestyle='--', alpha=0.5)

        self.frame_stability.fig.subplots_adjust(left=0.25)

        self.frame_stability.canvas.draw()

    def update_activity_view(self):
        """
        Обновляет вкладку «Деловая активность»: оборачиваемость и операционный цикл.
        """
        if not self.analysis_results:
            return

        inv_turn = self.analysis_results["Оборачиваемость запасов"]
        ar_turn = self.analysis_results["Оборачиваемость ДЗ"]
        ar_days = self.analysis_results["Период погашения ДЗ (дни)"]
        ap_turn = self.analysis_results["Оборачиваемость КЗ"]
        assets_turn = self.analysis_results["Оборачиваемость активов"]
        equity_turn = self.analysis_results["Оборачиваемость СК"]
        op_cycle = self.analysis_results["Операционный цикл"]

        activity_text = "ДЕЛОВАЯ АКТИВНОСТЬ:\n"
        activity_text += f"  Оборачиваемость запасов: {inv_turn:.2f} раз/год\n"
        activity_text += f"  Оборачиваемость ДЗ: {ar_turn:.2f} раз/год\n"
        activity_text += f"  Период погашения ДЗ: {ar_days:.0f} дней\n"
        activity_text += f"  Оборачиваемость КЗ: {ap_turn:.2f} раз/год\n"
        activity_text += f"  Оборачиваемость активов: {assets_turn:.2f} раз/год\n"
        activity_text += f"  Оборачиваемость СК: {equity_turn:.2f} раз/год\n"
        activity_text += f"  Операционный цикл: {op_cycle:.0f} дней\n\n"

        notes = []
        if ar_days > 90:
            notes.append("⚠ Срок погашения ДЗ > 90 дней — высокий риск просрочки.")
        if op_cycle > 120:
            notes.append("⚠ Операционный цикл > 120 дней — медленный оборот капитала.")
        if inv_turn < 4:
            notes.append("⚠ Запасы оборачиваются реже 4 раз в год — избыток складских запасов.")

        if notes:
            activity_text += "РЕКОМЕНДАЦИИ:\n" + "\n".join(f"  • {n}" for n in notes)

        self.frame_activity.text.config(state='normal')
        self.frame_activity.text.delete(1.0, tk.END)
        self.frame_activity.text.insert(tk.END, activity_text)
        self.frame_activity.text.config(state='disabled')

        self.frame_activity.ax.clear()
        labels = ["Запасы", "ДЗ", "КЗ", "Активы", "СК"]
        values = [inv_turn, ar_turn, ap_turn, assets_turn, equity_turn]
        colors = ['#4e79a7', '#f28e2b', '#e15759', '#59a14f', '#b07aa1']
        self.frame_activity.ax.bar(labels, values, color=colors)
        self.frame_activity.ax.set_ylabel("Обороты в год")
        self.frame_activity.ax.set_title("Оборачиваемость активов и обязательств")
        self.frame_activity.ax.grid(axis='y', linestyle='--', alpha=0.5)
        self.frame_activity.canvas.draw()

    # --- ВВОД/ВЫВОД ---
    def load_excel(self):
        """
        Загружает данные из Excel-файла с листом «Данные»,
        содержащим столбцы «Показатель» и «Значение».
        Заполняет соответствующие поля интерфейса.
        """
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл с данными",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            sheet = pd.read_excel(path, sheet_name="Данные")
            if "Показатель" not in sheet.columns or "Значение" not in sheet.columns:
                messagebox.showerror("Ошибка", "Файл должен содержать столбцы 'Показатель' и 'Значение'.")
                return

            sheet = sheet.dropna(subset=["Показатель"])
            data_map = dict(zip(sheet["Показатель"].astype(str).str.strip(), sheet["Значение"]))

            loaded_count = 0
            for field, entry in self.input_fields.items():
                if field in data_map:
                    val = data_map[field]
                    entry.delete(0, tk.END)
                    if pd.notna(val):
                        entry.insert(0, str(val))
                    loaded_count += 1

            if "Название предприятия" in data_map:
                val = data_map["Название предприятия"]
                self.company_name.delete(0, tk.END)
                if pd.notna(val):
                    self.company_name.insert(0, str(val))
                loaded_count += 1

            if "ИНН" in data_map:
                val = data_map["ИНН"]
                self.inn_field.delete(0, tk.END)
                if pd.notna(val):
                    self.inn_field.insert(0, str(val))
                loaded_count += 1

            messagebox.showinfo("Готово", f"Загружено {loaded_count} значений из файла.")
        except Exception as err:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{err}")

    def make_template(self):
        """
        Создаёт Excel-шаблон для заполнения данных.
        Содержит все необходимые показатели в первом столбце.
        """
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Шаблон_данных.xlsx"
        )
        if not path:
            return

        try:
            items = ["Название предприятия", "ИНН"] + list(self.input_fields.keys())
            template_df = pd.DataFrame({"Показатель": items, "Значение": [""] * len(items)})
            template_df.to_excel(path, index=False, sheet_name="Данные")
            messagebox.showinfo("Шаблон", f"Файл шаблона создан:\n{path}")
        except Exception as err:
            messagebox.showerror("Ошибка", f"Не удалось создать шаблон:\n{err}")

    def export_excel(self):
        """
        Экспортирует результаты анализа в Excel-файл.
        Включает название компании, ИНН и все рассчитанные коэффициенты.
        """
        if not self.analysis_results:
            messagebox.showwarning("Внимание", "Сначала выполните расчёт.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            output = {}
            name = self.company_name.get().strip()
            inn = self.inn_field.get().strip()
            if name:
                output["Компания"] = name
            if inn:
                output["ИНН"] = inn
            output.update(self.analysis_results)

            df = pd.DataFrame(list(output.items()), columns=["Показатель", "Значение"])
            df.to_excel(path, index=False, sheet_name="Результаты")
            messagebox.showinfo("Готово", f"Результаты сохранены в:\n{path}")
        except Exception as err:
            messagebox.showerror("Ошибка", f"Не удалось сохранить Excel:\n{err}")

    def export_pdf(self):
        """
        Экспортирует результаты в PDF-файл.
        Использует ReportLab, поддерживает кириллицу через шрифт DejaVuSans (если доступен).
        """
        if not self.analysis_results:
            messagebox.showwarning("Внимание", "Сначала выполните расчёт.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not path:
            return

        try:
            pdf = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            y_pos = h - 50

            try:
                font_file = "DejaVuSans.ttf"
                if os.path.exists(font_file):
                    pdfmetrics.registerFont(TTFont("DejaVu", font_file))
                    pdf.setFont("DejaVu", 12)
                else:
                    pdf.setFont("Helvetica", 12)
            except:
                pdf.setFont("Helvetica", 12)

            pdf.drawString(50, y_pos, "Финансовый анализ предприятия")
            y_pos -= 30

            name = self.company_name.get().strip()
            inn = self.inn_field.get().strip()
            if name:
                pdf.drawString(50, y_pos, f"Компания: {name}")
                y_pos -= 20
            if inn:
                pdf.drawString(50, y_pos, f"ИНН: {inn}")
                y_pos -= 20
            y_pos -= 10

            full_text = ""
            for tab_name, text_widget in [
                ("ЛИКВИДНОСТЬ", self.frame_liquidity.text),
                ("РЕТАБЕЛЬНОСТЬ", self.frame_profitability.text),
                ("ФИНАНСОВАЯ УСТОЙЧИВОСТЬ", self.frame_stability.text),
                ("ДЕЛОВАЯ АКТИВНОСТЬ", self.frame_activity.text)
            ]:
                full_text += f"\n{tab_name}\n" + "="*50 + "\n"
                full_text += text_widget.get(1.0, tk.END).strip() + "\n\n"

            for line in full_text.split('\n'):
                if y_pos < 50:
                    pdf.showPage()
                    y_pos = h - 50
                pdf.drawString(50, y_pos, line[:100])
                y_pos -= 15

            pdf.save()
            messagebox.showinfo("Готово", f"PDF-файл сохранён:\n{path}")
        except Exception as err:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл PDF:\n{err}")


if __name__ == "__main__":
    root = tk.Tk()
    # Включение DPI-awareness для Windows (более чёткое отображение на 4K-экранах)
    if os.name == 'nt':
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    app = FinancialAnalyzer(root)
    root.mainloop()

"""Interfaz gráfica para antecedentes judiciales - Policía Nacional.

El CAPTCHA se resuelve manualmente en la ventana de Chrome. La aplicación
lleva el resto: Excel, orden, llenado, consulta, PDF, estado y reportes.
"""

from __future__ import annotations

import json
import logging
import os
import platform
import re
from datetime import datetime
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import main as core
import sources
from config import Settings, chrome_path, load_settings, save_settings
from perfiles import CAMPOS

PERFILES_UI = ["TODOS", "AUXILIAR", "ENFERMERO", "MEDICO"]
BACKENDS = ["auto", "uc", "selenium"]

FILENAME_UI = {
    "APELLIDO NOMBRE - CEDULA": "apellido_primero",
    "CEDULA NOMBRE APELLIDO": "documento_primero",
}
FILENAME_INV = {value: key for key, value in FILENAME_UI.items()}

COLOR_OK = "#137333"
COLOR_WARN = "#b06000"
COLOR_ERROR = "#c5221f"


class QueueWriter:
    def __init__(self, target: queue.Queue):
        self.target = target

    def write(self, text):
        if text:
            self.target.put(("log", str(text)))

    def flush(self):
        pass

    def isatty(self):
        return False


class QueueLogHandler(logging.Handler):
    def __init__(self, target: queue.Queue):
        super().__init__()
        self.target = target
        self.setFormatter(logging.Formatter("%(asctime)s · %(levelname)s · %(message)s", "%H:%M:%S"))

    def emit(self, record):
        try:
            self.target.put(("log", self.format(record) + "\n"))
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.settings = load_settings()
        self._queue: queue.Queue = queue.Queue()
        self._worker: threading.Thread | None = None
        self._stop = threading.Event()
        self._old_stdout = None
        self._old_stderr = None
        self._display_to_original: dict[str, str] = {}
        self._columns_info: list[dict] = []
        self._loaded_signature = None
        self._loading = False

        self.title("Antecedentes Judiciales · Policía Nacional")
        self.geometry("1120x820")
        self.minsize(980, 700)

        self._build_style()
        self._build_vars()
        self._build_ui()

        self._load_settings_into_ui()
        self._install_log_handler()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(150, self._poll_queue)
        self.after(300, self._reload_source)

    # ------------------------------------------------------------------
    # Construcción
    # ------------------------------------------------------------------

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Helvetica", 17, "bold"))
        style.configure("Subtitle.TLabel", foreground="#5f6368")
        style.configure("Accent.TButton", font=("Helvetica", 11, "bold"))
        style.configure("Hint.TLabel", foreground="#5f6368")
        style.configure("Ok.TLabel", foreground=COLOR_OK)
        style.configure("Warn.TLabel", foreground=COLOR_WARN)

    def _build_vars(self):
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_row_var = tk.IntVar(value=1)
        self.map_vars: dict[str, tk.StringVar] = {campo.clave: tk.StringVar() for campo in CAMPOS}

        self.profile_var = tk.StringVar(value="TODOS")
        self.batch_var = tk.IntVar(value=0)
        self.resume_var = tk.BooleanVar(value=True)
        self.reset_var = tk.BooleanVar(value=False)
        self.attempts_var = tk.IntVar(value=2)
        self.manual_timeout_var = tk.IntVar(value=300)
        self.backend_var = tk.StringVar(value="auto")
        self.filename_var = tk.StringVar(value="APELLIDO NOMBRE - CEDULA")
        self.visible_var = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value="Listo.")
        self.progress_text_var = tk.StringVar(value="0 / 0   ·   OK 0   ·   Fallos 0")
        self.state_var = tk.StringVar(value="")
        self.excel_summary_var = tk.StringVar(value="Sin archivo válido.")
        self.chrome_var = tk.StringVar(value="")
        self.captcha_result_var = tk.StringVar(value="")

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=(16, 12, 16, 6))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Antecedentes Judiciales", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Policía Nacional de Colombia   |   consulta masiva con captcha manual",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w")
        ttk.Label(header, textvariable=self.chrome_var).grid(row=0, column=1, rowspan=2, sticky="e")

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=16, pady=(4, 8))

        self._build_tab_datos()
        self._build_tab_salida()
        self._build_tab_captcha()
        self._build_tab_ejecucion()

        self._build_log()
        self._build_footer()

    def _tab(self, title):
        frame = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(frame, text=title)
        frame.columnconfigure(0, weight=1)
        return frame

    # ---------------------------- Datos -------------------------------

    def _build_tab_datos(self):
        frame = self._tab("1 · Datos")

        file_box = ttk.LabelFrame(frame, text=" Archivo de entrada ", padding=10)
        file_box.grid(row=0, column=0, sticky="ew")
        file_box.columnconfigure(1, weight=1)

        ttk.Label(file_box, text="Excel de caracterización:").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(file_box, textvariable=self.excel_var)
        entry.grid(row=0, column=1, sticky="ew", padx=8)
        entry.bind("<Return>", lambda _e: self._reload_source())
        entry.bind("<FocusOut>", lambda _e: self._reload_source())
        ttk.Button(file_box, text="Examinar…", command=self._browse_excel).grid(row=0, column=2, padx=4)

        ttk.Label(file_box, text="Hoja:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.sheet_combo = ttk.Combobox(file_box, textvariable=self.sheet_var, state="readonly", width=34)
        self.sheet_combo.grid(row=1, column=1, sticky="w", padx=8, pady=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._reload_source())

        header_frame = ttk.Frame(file_box)
        header_frame.grid(row=1, column=2, sticky="e", pady=(8, 0))
        ttk.Label(header_frame, text="Encabezados en fila:").pack(side="left")
        spin = ttk.Spinbox(
            header_frame, from_=1, to=100, width=5, textvariable=self.header_row_var,
            command=self._reload_source,
        )
        spin.pack(side="left", padx=6)
        spin.bind("<Return>", lambda _e: self._reload_source())
        spin.bind("<FocusOut>", lambda _e: self._reload_source())

        mapping_box = ttk.LabelFrame(frame, text=" Mapeo de columnas ", padding=8)
        mapping_box.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        mapping_box.columnconfigure(1, weight=1)
        mapping_box.columnconfigure(3, weight=1)

        for index, campo in enumerate(CAMPOS):
            row, col = divmod(index, 2)
            ttk.Label(mapping_box, text=campo.etiqueta + ":").grid(
                row=row, column=col * 2, sticky="w", pady=2
            )
            combo = ttk.Combobox(
                mapping_box, textvariable=self.map_vars[campo.clave], state="readonly", width=28
            )
            combo.grid(row=row, column=col * 2 + 1, sticky="w", padx=(8, 18), pady=2)
            self.__dict__[f"combo_{campo.clave}"] = combo

        ttk.Button(
            mapping_box, text="Detectar automáticamente", command=self._apply_detected_mapping
        ).grid(row=(len(CAMPOS) + 1) // 2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        preview_box = ttk.LabelFrame(frame, text=" Vista previa ", padding=10)
        preview_box.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        preview_box.columnconfigure(0, weight=1)
        preview_box.rowconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        self.preview = ttk.Treeview(preview_box, show="headings", height=4)
        self.preview.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(preview_box, orient="vertical", command=self.preview.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(preview_box, orient="horizontal", command=self.preview.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.preview.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        bottom = ttk.Frame(frame)
        bottom.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.excel_summary_var, style="Hint.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(bottom, text="Validar datos", command=self._validate).grid(row=0, column=1, padx=4)

    # ---------------------------- Salida ------------------------------

    def _build_tab_salida(self):
        frame = self._tab("2 · Salida y navegador")

        box = ttk.LabelFrame(frame, text=" Carpetas ", padding=10)
        box.grid(row=0, column=0, sticky="ew")
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="Carpeta de salida:").grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.output_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(box, text="Examinar…", command=self._browse_output).grid(row=0, column=2, padx=4)
        ttk.Button(box, text="Abrir", command=lambda: self._open_path(self.output_var.get())).grid(
            row=0, column=3
        )

        ttk.Label(box, text="Nombre del PDF:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            box, textvariable=self.filename_var, values=list(FILENAME_UI), state="readonly", width=32
        ).grid(row=1, column=1, sticky="w", padx=8, pady=(8, 0))

        ttk.Label(box, text="Dentro de la carpeta se crean AUXILIAR, ENFERMERO y MEDICO.",
                  style="Hint.TLabel").grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        browser = ttk.LabelFrame(frame, text=" Navegador ", padding=10)
        browser.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        browser.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            browser,
            text="Mostrar la ventana de Chrome (obligatorio para resolver el captcha)",
            variable=self.visible_var,
            state="disabled",
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(browser, text="Driver:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            browser, textvariable=self.backend_var, values=BACKENDS, state="readonly", width=18
        ).grid(row=1, column=1, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(
            browser,
            text="auto = undetected-chromedriver y, si falla, Selenium estándar.",
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(
            actions, text="Abrir carpeta de salida", command=lambda: self._open_path(self.output_var.get())
        ).pack(side="left")
        ttk.Button(actions, text="Actualizar estado", command=lambda: self._refresh_state(reconcile=True)).pack(side="left", padx=8)

    # ---------------------------- CAPTCHA -----------------------------

    def _build_tab_captcha(self):
        frame = self._tab("3 · CAPTCHA manual")

        info = ttk.LabelFrame(frame, text=" Cómo funciona ", padding=12)
        info.grid(row=0, column=0, sticky="ew")

        ttk.Label(
            info,
            text=(
                "El sitio de la Policía usa reCAPTCHA v2 de Google: un checkbox y, si Google lo pide, "
                "una cuadrícula de imágenes (por ejemplo, «selecciona todas las imágenes de autobús»).\n\n"
                "1. La aplicación abre Chrome y llena la cédula.\n"
                "2. Si aparece la cuadrícula, selecciona las imágenes y pulsa «Verificar».\n"
                "3. La aplicación detecta el token y pulsa «Consultar» automáticamente.\n\n"
                "No necesitas API keys ni pagar 2captcha. Si el checkbox pasa solo, mejor."
            ),
            wraplength=900,
            justify="left",
        ).grid(row=0, column=0, sticky="w")

        timeout_box = ttk.LabelFrame(frame, text=" Tiempo de espera ", padding=10)
        timeout_box.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(timeout_box, text="Esperar el captcha manual hasta (segundos):").pack(side="left")
        ttk.Spinbox(
            timeout_box, from_=60, to=900, width=8, textvariable=self.manual_timeout_var
        ).pack(side="left", padx=8)
        ttk.Label(
            timeout_box,
            text="Si se agota, esa persona queda en el reporte y puedes reintentarla después.",
            style="Hint.TLabel",
        ).pack(side="left", padx=8)

        test_box = ttk.LabelFrame(frame, text=" Prueba ", padding=10)
        test_box.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(test_box, text="Probar CAPTCHA (abre Chrome)", command=self._test_captcha).pack(
            side="left"
        )
        ttk.Label(test_box, textvariable=self.captcha_result_var, style="Hint.TLabel").pack(
            side="left", padx=8
        )

        ttk.Label(
            frame,
            text="La prueba solo abre el formulario y espera el token; no consulta ninguna cédula.",
            style="Hint.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))

    # --------------------------- Ejecución ----------------------------

    def _build_tab_ejecucion(self):
        frame = self._tab("4 · Ejecución")

        run_box = ttk.LabelFrame(frame, text=" Lote ", padding=10)
        run_box.grid(row=0, column=0, sticky="ew")
        run_box.columnconfigure(1, weight=1)

        ttk.Label(run_box, text="Perfil:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            run_box, textvariable=self.profile_var, values=PERFILES_UI, state="readonly", width=18
        ).grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(run_box, text="Cantidad (0 = todos):").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Spinbox(run_box, from_=0, to=10000, textvariable=self.batch_var, width=8).grid(
            row=1, column=1, sticky="w", padx=8, pady=(6, 0)
        )

        ttk.Checkbutton(run_box, text="Reanudar (omitir procesados)", variable=self.resume_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            run_box, text="Reiniciar estado (procesar todo de nuevo)", variable=self.reset_var
        ).grid(row=3, column=0, columnspan=2, sticky="w")

        tuning = ttk.LabelFrame(frame, text=" Reintentos ", padding=10)
        tuning.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(tuning, text="Intentos por persona si el captcha es rechazado:").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Spinbox(tuning, from_=1, to=5, textvariable=self.attempts_var, width=6).grid(
            row=0, column=1, sticky="w", padx=8
        )

        state_box = ttk.LabelFrame(frame, text=" Estado actual ", padding=10)
        state_box.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        state_box.columnconfigure(0, weight=1)
        ttk.Label(state_box, textvariable=self.state_var, style="Hint.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(state_box, text="Actualizar", command=lambda: self._refresh_state(reconcile=True)).grid(row=0, column=1)

        ttk.Label(
            frame,
            text="El orden de procesamiento es exactamente el de las filas del Excel.",
            style="Hint.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))

    # ----------------------------- Log --------------------------------

    def _build_log(self):
        box = ttk.LabelFrame(self, text=" Registro ", padding=(10, 6))
        box.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 8))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            box, height=7, wrap="word", font=("Courier New", 10), state="disabled",
            background="#101418", foreground="#e6edf3", insertbackground="#e6edf3",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(box, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        self.log_text.tag_configure("info", foreground="#e6edf3")
        self.log_text.tag_configure("ok", foreground="#7ee787")
        self.log_text.tag_configure("warn", foreground="#ffd166")
        self.log_text.tag_configure("error", foreground="#ff7b72")

    def _build_footer(self):
        footer = ttk.Frame(self, padding=(16, 0, 16, 14))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(footer, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        ttk.Label(footer, textvariable=self.progress_text_var).grid(row=0, column=1, padx=(0, 12))

        self.start_button = ttk.Button(
            footer, text="Iniciar consulta", style="Accent.TButton", command=self._start
        )
        self.start_button.grid(row=0, column=2, padx=4)
        self.stop_button = ttk.Button(footer, text="Detener", command=self._request_stop, state="disabled")
        self.stop_button.grid(row=0, column=3, padx=4)

        ttk.Label(footer, textvariable=self.status_var, style="Hint.TLabel").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(6, 0)
        )

    # ------------------------------------------------------------------
    # Configuración -> interfaz
    # ------------------------------------------------------------------

    def _load_settings_into_ui(self):
        settings = self.settings
        self.excel_var.set(str(settings.excel_path))
        self.output_var.set(str(settings.output_dir))
        self.header_row_var.set(int(settings.header_row or 1))
        self.attempts_var.set(int(settings.captcha_max_attempts or 2))
        self.manual_timeout_var.set(int(settings.manual_timeout or 300))
        self.backend_var.set(settings.driver_backend or "auto")
        self.resume_var.set(bool(settings.resume))
        self.profile_var.set(settings.profiles[0] if settings.profiles else "TODOS")
        self.batch_var.set(int(settings.batch_size or 0))
        self.filename_var.set(FILENAME_INV.get(settings.filename_pattern, "APELLIDO NOMBRE - CEDULA"))
        self.visible_var.set(True)

        if settings.sheet_name not in (None, ""):
            self.sheet_var.set(str(settings.sheet_name))

        chrome = chrome_path()
        if chrome:
            self.chrome_var.set(f"Chrome: OK  ·  {chrome}")
        else:
            self.chrome_var.set("Chrome: NO ENCONTRADO")

        self._refresh_state()

    def _from_ui(self) -> Settings:
        path_text = self.excel_var.get().strip()
        if path_text and Path(path_text).exists():
            signature = (path_text, self.sheet_var.get(), int(self.header_row_var.get() or 1))
            if signature != self._loaded_signature:
                self._reload_source()

        settings = Settings()
        settings.excel_path = Path(self.excel_var.get().strip() or ".")
        settings.output_dir = Path(self.output_var.get().strip() or ".")
        settings.sheet_name = self.sheet_var.get().strip() or None
        settings.header_row = max(1, int(self.header_row_var.get() or 1))
        settings.column_map = {
            clave: self._display_to_original[var.get()]
            for clave, var in self.map_vars.items()
            if var.get() and var.get() in self._display_to_original
        }
        settings.profiles = None if self.profile_var.get() == "TODOS" else [self.profile_var.get()]
        settings.batch_size = max(0, int(self.batch_var.get() or 0))
        settings.headless = False
        settings.resume = bool(self.resume_var.get())
        settings.reset_state = bool(self.reset_var.get())
        settings.captcha_max_attempts = max(1, int(self.attempts_var.get() or 2))
        settings.manual_timeout = max(60, int(self.manual_timeout_var.get() or 300))
        settings.driver_backend = self.backend_var.get() or "auto"
        settings.filename_pattern = FILENAME_UI.get(
            self.filename_var.get(), "apellido_primero"
        )
        settings.interactive = False
        return settings

    def _save_config(self, silent=False):
        try:
            self.settings = self._from_ui()
            save_settings(self.settings)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Configuración", f"No se pudo guardar config.json:\n{exc}")

    # ------------------------------------------------------------------
    # Lectura del Excel
    # ------------------------------------------------------------------

    def _browse_excel(self):
        initial = str(Path(self.excel_var.get()).parent) if self.excel_var.get() else str(Path.home())
        path = filedialog.askopenfilename(
            title="Selecciona la caracterización",
            initialdir=initial,
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")],
        )
        if path:
            self.excel_var.set(path)
            self._reload_source(force=True)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Carpeta de salida")
        if path:
            self.output_var.set(path)

    def _reload_source(self, force: bool = False):
        if self._loading:
            return
        path = Path(self.excel_var.get().strip())
        if not path.exists():
            self._clear_preview()
            return

        signature = (str(path), self.sheet_var.get(), int(self.header_row_var.get() or 1))
        if not force and signature == self._loaded_signature and self._columns_info:
            return

        self._loading = True
        try:
            sheets = sources.list_sheets(path)
            self.sheet_combo.configure(values=sheets)
            current = self.sheet_var.get()
            if current not in sheets:
                preferred = str(self.settings.sheet_name or "")
                current = preferred if preferred in sheets else sheets[0]
                self.sheet_var.set(current)

            info = sources.inspect(
                path, sheet_name=self.sheet_var.get() or None,
                header_row=int(self.header_row_var.get() or 1),
            )
            self._columns_info = info["columnas"]
            self._display_to_original = {c["display"]: c["original"] for c in info["columnas"]}
            displays = [""] + list(self._display_to_original)

            for campo in CAMPOS:
                combo = self.__dict__.get(f"combo_{campo.clave}")
                if combo is not None:
                    combo.configure(values=displays)

            self._apply_detected_mapping(info["mapeo_detectado"], silent=True)
            self._populate_preview(info)
            self.excel_summary_var.set(
                f"Hoja '{info['hoja']}' · {info['total_filas']} filas en la hoja · "
                f"encabezados en la fila {info['fila_encabezado']}"
            )
            self._loaded_signature = (str(path), info["hoja"], int(self.header_row_var.get() or 1))
        except Exception as exc:
            self._clear_preview()
            self._log(f"[ERROR] No se pudo leer el Excel: {exc}\n")
        finally:
            self._loading = False

    def _apply_detected_mapping(self, mapping=None, silent=False):
        if mapping is None:
            mapping = sources.detect_mapping([c["original"] for c in self._columns_info])

        original_to_display = {c["original"]: c["display"] for c in self._columns_info}
        for campo in CAMPOS:
            if not campo.obligatorio:
                self.map_vars[campo.clave].set("")
        for clave, original in mapping.items():
            var = self.map_vars.get(clave)
            if var is not None and original in original_to_display:
                var.set(original_to_display[original])

        if not silent:
            self._log("[INFO] Mapeo de columnas detectado automáticamente.\n")

    def _populate_preview(self, info):
        original = info["columnas"]
        by_display = {c["display"]: c for c in original}

        order = [self.map_vars[campo.clave].get() for campo in CAMPOS]
        order = [value for value in order if value and value in by_display]
        ordered = [by_display[value] for value in order]
        ordered += [c for c in original if c["display"] not in order]

        # Si se reordenan los encabezados, hay que reordenar también los valores.
        index_by_display = {c["display"]: i for i, c in enumerate(original)}
        indices = [index_by_display[c["display"]] for c in ordered]

        # Etiqueta clara: nombre del campo y columna real del Excel.
        labels = {}
        for campo in CAMPOS:
            display = self.map_vars[campo.clave].get()
            if display and display in by_display:
                etiqueta = campo.etiqueta.replace(" (opcional)", "")
                labels[display] = f"{etiqueta} · {by_display[display]['letra']}"
        for column in original:
            labels.setdefault(column["display"], column["display"])

        ids = [f"c{i}" for i in range(len(ordered))]
        self.preview.delete(*self.preview.get_children())
        self.preview.configure(columns=ids, show="headings")

        for index, column in enumerate(ordered):
            self.preview.heading(ids[index], text=labels[column["display"]])
            self.preview.column(ids[index], width=135, minwidth=90, anchor="w", stretch=True)

        for row in info["vista_previa"]:
            values = [row[i] if i < len(row) else "" for i in indices]
            self.preview.insert("", "end", values=values)

    def _clear_preview(self):
        self.preview.delete(*self.preview.get_children())
        self.preview.configure(columns=(), show="headings")
        self._columns_info = []
        self._display_to_original = {}
        self.excel_summary_var.set("Sin archivo válido.")

    def _validate(self):
        self._save_config(silent=True)
        settings = self._from_ui()
        try:
            users, incomplete, stats = sources.load_users(
                settings.excel_path,
                sheet_name=settings.sheet_name,
                header_row=settings.header_row,
                column_map=settings.column_map,
                selected_profiles=settings.profiles,
            )
        except Exception as exc:
            messagebox.showerror("Validación", str(exc))
            self._log(f"[ERROR] Validación: {exc}\n")
            return

        self.excel_summary_var.set(
            f"{stats['validos']} válidos · {stats['incompletos']} incompletos · "
            f"{stats['duplicados']} duplicados · {stats['filtrados_por_perfil']} filtrados"
        )
        self._log(
            f"[OK] Validación: {stats['validos']} válidos, {stats['incompletos']} incompletos, "
            f"{stats['duplicados']} duplicados.\n"
        )
        for fila in incomplete[:10]:
            self._log(f"[WARN] fila {fila['fila']}: {fila['detalle']}\n")
        for fila in stats["duplicados_detalle"][:10]:
            self._log(
                f"[WARN] fila {fila['fila']}: documento {fila['documento']} repetido "
                f"(primera vez en la fila {fila['primera_fila']})\n"
            )

        messagebox.showinfo(
            "Validación",
            f"Hoja: {stats['hoja']}\n"
            f"Válidos: {stats['validos']}\n"
            f"Incompletos: {stats['incompletos']}\n"
            f"Duplicados: {stats['duplicados']}\n"
            f"Filtrados por perfil: {stats['filtrados_por_perfil']}",
        )

    # ------------------------------------------------------------------
    # Estado
    # ------------------------------------------------------------------

    def _reconcile_state_with_files(self) -> int:
        """Quita de 'procesados' las cédulas cuyo PDF ya no existe."""
        path = Path(self.settings.state_file)
        if not path.exists():
            return 0

        output = Path(self.output_var.get().strip())
        if not output.exists():
            self._log("[WARN] La carpeta de salida no existe; no se sincroniza el estado.\n")
            return 0

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return 0

        processed = [str(doc) for doc in data.get("procesados", [])]
        if not processed:
            return 0

        names = [p.name for p in output.glob("*.pdf")]
        names += [p.name for p in output.glob("*/*.pdf")]

        missing = []
        for doc in processed:
            pattern = re.compile(rf"(?<!\d){re.escape(doc)}(?!\d)")
            if not any(pattern.search(name) for name in names):
                missing.append(doc)

        if not missing:
            return 0

        missing_set = set(missing)
        data["procesados"] = [doc for doc in processed if doc not in missing_set]
        data["ultima_actualizacion"] = datetime.now().isoformat(timespec="seconds")

        tmp = path.with_name(path.name + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, path)
        return len(missing)

    def _refresh_state(self, reconcile: bool = False):
        path = Path(self.settings.state_file)
        if not path.exists():
            self.state_var.set("Sin estado previo (primera corrida).")
            return

        if reconcile:
            removed = self._reconcile_state_with_files()
            if removed:
                self._log(
                    f"[WARN] Se quitaron {removed} procesados porque su PDF ya no está en la carpeta.\n"
                )

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            procesados = len(data.get("procesados", []))
            fallidos = len(data.get("fallidos", []))
            self.state_var.set(f"Procesados: {procesados} · Fallidos: {fallidos} · {path.name}")
        except Exception as exc:
            self.state_var.set(f"No se pudo leer el estado: {exc}")

    # ------------------------------------------------------------------
    # Prueba de CAPTCHA
    # ------------------------------------------------------------------

    def _test_captcha(self):
        settings = self._from_ui()
        self.captcha_result_var.set("Abriendo Chrome…")
        self._log("[INFO] Abriendo Chrome para probar el captcha…\n")

        def worker():
            driver = None
            try:
                core.apply_settings(settings)
                driver = core.create_driver(headless=False, backend=settings.driver_backend)
                driver.maximize_window()
                core.focus_chrome()
                wait = core.WebDriverWait(driver, 30)
                core.ensure_form(driver, wait)

                self._queue.put(("status", "Resuelve el captcha en Chrome…"))
                if core.try_checkbox(driver, timeout=6):
                    self._queue.put(("captcha_ok", "Checkbox automático (token listo)"))
                    return

                ok = core.wait_manual_token(
                    driver, settings.manual_timeout, "__prueba__",
                    on_status=lambda text: self._queue.put(("status", text)),
                )
                if ok:
                    self._queue.put(("captcha_ok", "Token listo (manual)"))
                else:
                    self._queue.put(("captcha_error", "No se detectó el token del captcha."))
            except Exception as exc:
                self._queue.put(("captcha_error", f"{type(exc).__name__}: {exc}"))
            finally:
                if driver:
                    try:
                        driver.quit()
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Ejecución
    # ------------------------------------------------------------------

    def _start(self):
        if self._worker and self._worker.is_alive():
            return

        settings = self._from_ui()
        if not Path(settings.excel_path).exists():
            messagebox.showerror("Falta el Excel", "Selecciona un archivo de caracterización válido.")
            return
        if not self.output_var.get().strip():
            messagebox.showerror("Falta la salida", "Selecciona la carpeta de salida.")
            return
        if settings.reset_state and not messagebox.askyesno(
            "Reiniciar estado",
            "Se consultarán TODOS los registros otra vez, incluidos los ya descargados.\n¿Continuar?",
        ):
            return

        reset_state = bool(settings.reset_state)
        self._save_config(silent=True)
        settings = self._from_ui()
        settings.reset_state = reset_state
        self.reset_var.set(False)

        try:
            Path(settings.output_dir).mkdir(parents=True, exist_ok=True)
            Path(settings.error_dir).mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Carpetas", f"No se pudieron crear las carpetas:\n{exc}")
            return

        self._log("\n" + "=" * 70 + "\n")
        self._log("[INFO] Iniciando consulta…\n")
        self._log("[INFO] Chrome se abrirá visible. Resuelve el captcha cuando aparezca.\n")
        self.status_var.set("Ejecutando…")
        self.progress.configure(maximum=100, value=0)
        self.progress_text_var.set("0 / 0   ·   OK 0   ·   Fallos 0")

        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self._stop.clear()

        self._old_stdout, self._old_stderr = sys.stdout, sys.stderr
        sys.stdout = QueueWriter(self._queue)
        sys.stderr = sys.stdout

        def worker():
            try:
                result = core.run(
                    settings,
                    on_progress=lambda done, total, ok, fail: self._queue.put(
                        ("progress", done, total, ok, fail)
                    ),
                    on_status=lambda text: self._queue.put(("status", text)),
                    should_stop=self._stop.is_set,
                )
                self._queue.put(("done", result))
            except Exception:
                self._queue.put(("crashed", traceback.format_exc()))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _request_stop(self):
        self._stop.set()
        self.status_var.set("Deteniendo… (termina la persona en curso)")
        self._log("[WARN] Detención solicitada. Se cerrará al terminar el registro actual.\n")
        self.stop_button.configure(state="disabled")

    def _finish_job(self):
        if self._old_stdout is not None:
            sys.stdout = self._old_stdout
            sys.stderr = self._old_stderr
            self._old_stdout = self._old_stderr = None

        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.status_var.set("Listo.")
        self._refresh_state()

    # ------------------------------------------------------------------
    # Cola / log
    # ------------------------------------------------------------------

    def _install_log_handler(self):
        handler = QueueLogHandler(self._queue)
        handler.setLevel(logging.INFO)
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)
        logging.getLogger("httpx").setLevel(logging.WARNING)
        logging.getLogger("selenium").setLevel(logging.WARNING)
        logging.getLogger("urllib3").setLevel(logging.WARNING)

    def _poll_queue(self):
        while True:
            try:
                item = self._queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "log":
                self._log(item[1])
            elif kind == "progress":
                _, done, total, ok, fail = item
                self.progress.configure(maximum=max(1, total), value=done)
                self.progress_text_var.set(f"{done} / {total}   ·   OK {ok}   ·   Fallos {fail}")
            elif kind == "status":
                self.status_var.set(item[1])
            elif kind == "captcha_ok":
                self.captcha_result_var.set(item[1])
                self._log(f"[OK] CAPTCHA: {item[1]}\n")
                messagebox.showinfo("Prueba de CAPTCHA", item[1])
            elif kind == "captcha_error":
                self.captcha_result_var.set("Error: " + item[1])
                self._log(f"[ERROR] CAPTCHA: {item[1]}\n")
                messagebox.showerror("Prueba de CAPTCHA", item[1])
            elif kind == "done":
                result = item[1]
                self._finish_job()
                self._log(
                    f"[OK] Terminado: {result.get('success', 0)} exitosos, "
                    f"{result.get('failed', 0)} fallidos.\n"
                )
                messagebox.showinfo(
                    "Consulta terminada",
                    f"Exitosos: {result.get('success', 0)}\n"
                    f"Fallidos: {result.get('failed', 0)}\n"
                    f"Carpeta: {result.get('carpeta', '')}",
                )
            elif kind == "crashed":
                self._finish_job()
                self._log("[ERROR] Fallo inesperado:\n" + item[1] + "\n")
                messagebox.showerror("Error", item[1][-1200:])

        self.after(150, self._poll_queue)

    def _log(self, text: str):
        if not text:
            return
        upper = text.upper()
        if "[ERROR]" in upper or "ERROR GENERAL" in upper or "TRACEBACK" in upper:
            tag = "error"
        elif "[WARN]" in upper or "[SKIP]" in upper:
            tag = "warn"
        elif "[OK]" in upper:
            tag = "ok"
        else:
            tag = "info"

        self.log_text.configure(state="normal")
        self.log_text.insert("end", text, tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------------
    # Utilidades
    # ------------------------------------------------------------------

    def _open_path(self, value: str):
        path = Path(value).expanduser()
        if not path.exists():
            messagebox.showwarning("Carpeta", f"No existe:\n{path}")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Abrir carpeta", str(exc))

    def _on_close(self):
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno("Salir", "Hay una consulta en curso. ¿Detenerla y salir?"):
                return
            self._stop.set()

        self._save_config(silent=True)
        self.destroy()


def main():
    try:
        if platform.system() == "Windows":
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

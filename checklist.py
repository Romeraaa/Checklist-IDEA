#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Checklist Onboarding IDEA TSG - Desktop App (Tkinter)
v7.2:
- Se abre maximizada (state('zoomed')).
- Interfaz adaptable y lista de tareas que crece dinámicamente.
- Fuentes Montserrat 11 pt embebidas (sin instalar).
- Logo, rojo IDEA, combobox Técnico, CSV export, autoguardado, scroll.
- FIX: Inicialización temprana de _task_vars/_note_vars para evitar AttributeError.
- Métodos de progreso robustos (funcionan antes de crear los checkboxes).
"""
import json, os, re, sys, ctypes, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import subprocess

APP_TITLE = "Checklist Preparación de Equipos · IDEA TSG"
STATE_FILE = "state.json"
RED = "#e30613"
BG = "#f6f7f9"
CARD_BG = "#ffffff"

BASE_TASKS = [
    "CREACIÓN DE USUARIOS AD/365","PERMISOS AD","LICENCIAS, PUESTO Y OFICINA 365",
    "ACCESOS DIRECTOS (EXACT, 7-ZIP, VPN ETC)","INSTALAR PROGRAMAS CONCRETOS","CONFIGURAR VPN",
    "INSTALAR MS TEAMS","ADOBE, 7-ZIP Y CHROME PREDETERMINADOS","ORDENAR BARRA DE TAREAS",
    "INSTALAR IMPRESORA Y PONER EN BLANCO Y NEGRO","SOLICITAR CLAVES WIN/OFFICE","ACTIVAR CLAVES WIN/OFFICE",
    "INICIAR SESION OUTLOOK/TEAMS","CONFIGURAR FUENTE OPEN SANS","CONFIGURAR FIRMA CORPORATIVA (borrar correo)",
    "SI TIENE 2 CORREOS: DESACTIVAR LEIDO / CONFIGURAR FIRMA","CONFIGURAR FIRMA EN OUTLOOK WEB",
    "ENVIAR CORREO RRHH CREDS EXACT/CORREO","BUSCAR ACTUALIZACIONES","POSIT CON CONTRASEÑA","MONTAR PUESTO",
    "COMPROBACIONES INVENTARIO",
    "CREAR ACCESO DIRECTO ONEDRIVE"
]
TECHNICIANS = ["Adrián Romera", "Javier Hernández", "Mario Aniorte"]

def resource_path(rel):
    """Soporta PyInstaller (sys._MEIPASS) y ejecución normal."""
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, rel)

def add_font(ttf_path):
    """Registrar fuente TTF en memoria (Windows) para no exigir instalación."""
    if os.name == "nt" and os.path.exists(ttf_path):
        FR_PRIVATE, FR_NOT_ENUM = 0x10, 0x20
        ctypes.windll.gdi32.AddFontResourceExW(os.path.abspath(ttf_path), FR_PRIVATE | FR_NOT_ENUM, 0)

def default_state():
    return {
        "joinerName": "", "technician": "", "requestNumber": "",
        "date": datetime.now().strftime("%Y-%m-%d"),
        "tasks": [{"name": t, "done": False, "notes": ""} for t in BASE_TASKS],
    }

def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        names = [t["name"] for t in data.get("tasks", [])]
        if names != BASE_TASKS:
            lookup = {t["name"]: t for t in data.get("tasks", [])}
            data["tasks"] = [{"name": n, "done": bool(lookup.get(n, {}).get("done", False)),
                              "notes": str(lookup.get(n, {}).get("notes", ""))} for n in BASE_TASKS]
        data.setdefault("joinerName", "")
        data.setdefault("technician", "")
        data.setdefault("requestNumber", "")
        data.setdefault("date", datetime.now().strftime("%Y-%m-%d"))
        return data
    except Exception:
        return default_state()

def save_state(state):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

class App(tk.Tk):

    def _generate_material(self):
        """Open a new CMD, pipe the joiner name to the EXE, and keep the checklist alive."""
        import tempfile

        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning(
                "Nombre requerido",
                "Indica el nombre de la incorporacion antes de generar el documento."
            )
            return

        exe_path = resource_path(r"\\172.16.0.220\Control_Informatica\3_Inventario\4_Automatizaciones\Documento_Material\Script_Documento_Material Rev-1.2.exe")  # ajusta la ruta si esta en otro sitio
        if not os.path.exists(exe_path):
            messagebox.showerror("Error", f"No se encontro el ejecutable:\n{exe_path}")
            return

        exe_dir = os.path.dirname(exe_path)

        # Escapar caracteres especiales de CMD (& | < > ^)
        def _escape_cmd(s: str) -> str:
            for ch in "^&|<>":
                s = s.replace(ch, "^" + ch)
            return s

        safe_name = _escape_cmd(name)

        # .bat temporal: cambia a la carpeta del exe, fuerza UTF-8 y le inyecta el nombre por stdin
        bat_lines = [
            "@echo off",
            "chcp 65001 >nul",
            f'pushd "{exe_dir}"',
            f'echo {safe_name} | "{exe_path}"',
            "popd"
        ]
        bat_path = os.path.join(tempfile.gettempdir(), "run_material_temp.bat")
        with open(bat_path, "w", encoding="utf-8", newline="") as f:
            f.write("\r\n".join(bat_lines))

        try:
            # Nueva consola independiente que ejecuta el .bat y se cierra sola
            subprocess.Popen(
                ["cmd.exe", "/c", "start", "", "cmd.exe", "/c", bat_path],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            messagebox.showinfo(
                "Ejecutando",
                f"Se abrio una consola independiente para generar el documento de:\n{name}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar el programa:\n{e}")


    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.configure(bg=BG)

        # Abrir maximizada (con barra de título)
        try:
            self.state('zoomed')
        except Exception:
            try: self.attributes('-zoomed', True)
            except Exception: pass

        # Icono de la ventana (opcional)
        try:
            self.iconbitmap(resource_path("idea.ico"))
        except Exception:
            pass

        # Cargar fuentes Montserrat (embebidas, opcionales)
        add_font(resource_path("Montserrat-Regular.ttf"))
        add_font(resource_path("Montserrat-Bold.ttf"))

        # Estilos
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except: pass
        try:
            style.configure(".", font=("Montserrat", 11))
            style.configure("Title.TLabel", foreground=RED, font=("Montserrat", 22, "bold"))
        except:
            # fallback si no cargan las fuentes
            style.configure(".", font=("Segoe UI", 10))
            style.configure("Title.TLabel", foreground=RED, font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", foreground="#6b7280", font=("Segoe UI", 10))
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure("Primary.TButton", padding=10, foreground="white", background=RED, font=("Montserrat", 11))
        style.map("Primary.TButton", background=[("active", RED), ("disabled", "#c4c4c4")])
        style.configure("TProgressbar", troughcolor="#e5e7eb", background=RED)

        # Estado guardado
        self.state_data = load_state()

        # 🔧 FIX: Inicializar estructuras ANTES de calcular progreso
        self._task_vars = []
        self._note_vars = []

        # Contenedor raíz expandible
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        card = ttk.Frame(outer, style="Card.TFrame", padding=20)
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        # fila 3 contendrá la lista y crecerá
        card.rowconfigure(3, weight=1)

        # Header con logo
        header = tk.Frame(card, bg=CARD_BG)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.logo_img = None
        lp = resource_path("logo.png")
        if os.path.exists(lp):
            try:
                img = tk.PhotoImage(file=lp)
                factor = max(1, img.width() // 48)
                self.logo_img = img.subsample(factor, factor)
            except Exception:
                self.logo_img = None
        if self.logo_img:
            tk.Label(header, image=self.logo_img, bg=CARD_BG).pack(side="left", padx=(4, 12))
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="Basado en la checklist provisional REV. 01",
                  style="Sub.TLabel").pack(side="left", padx=16)

        # Formulario superior
        form = tk.Frame(card, bg=CARD_BG)
        form.grid(row=1, column=0, sticky="ew", pady=(8, 12))
        for i in range(4):
            form.columnconfigure(i, weight=1)

        ttk.Label(form, text="Nombre incorporación:").grid(row=0, column=0, sticky="w")
        self.name_var = tk.StringVar(value=self.state_data["joinerName"])
        ttk.Entry(form, textvariable=self.name_var).grid(row=0, column=1, sticky="ew", padx=(6, 20))

        ttk.Label(form, text="Fecha:").grid(row=0, column=2, sticky="w")
        self.date_var = tk.StringVar(value=self.state_data["date"])
        ttk.Entry(form, textvariable=self.date_var, width=16).grid(row=0, column=3, sticky="ew", padx=(6, 20))

        ttk.Label(form, text="Técnico:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.tech_var = tk.StringVar(value=self.state_data.get("technician", ""))
        self.tech_combo = ttk.Combobox(form, textvariable=self.tech_var,
                                       values=TECHNICIANS, state="readonly")
        self.tech_combo.grid(row=1, column=1, sticky="ew", padx=(6, 20), pady=(8, 0))
        self.tech_combo.bind("<<ComboboxSelected>>", lambda e: self._autosave())

        ttk.Label(form, text="Número de petición:").grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.req_var = tk.StringVar(value=self.state_data.get("requestNumber", ""))
        ttk.Entry(form, textvariable=self.req_var, width=16).grid(row=1, column=3, sticky="ew", padx=(6, 20), pady=(8, 0))

        # Progreso (ya puede usar _progress_pct de forma robusta)
        pb_frame = tk.Frame(card, bg=CARD_BG)
        pb_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.progress_var = tk.IntVar(value=self._progress_pct())
        self.progress = ttk.Progressbar(pb_frame, orient="horizontal", mode="determinate",
                                        maximum=100, variable=self.progress_var)
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.progress_label = ttk.Label(pb_frame, text=self._progress_text())
        self.progress_label.pack(side="left")

        # Lista de tareas expandible
        list_frame = ttk.Frame(card, style="Card.TFrame", padding=10)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(list_frame, highlightthickness=0, bg=CARD_BG)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.task_container = tk.Frame(self.canvas, bg=CARD_BG)
        self.task_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.task_container, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Rueda del ratón
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)      # Win/mac
        self.canvas.bind_all("<Button-4>", lambda e: self._on_mousewheel_linux(-1))  # Linux up
        self.canvas.bind_all("<Button-5>", lambda e: self._on_mousewheel_linux(1))   # Linux down

        # Tareas (rellenan _task_vars y _note_vars)
        for idx, task in enumerate(self.state_data["tasks"]):
            row = tk.Frame(self.task_container, bg=CARD_BG)
            row.grid(row=idx, column=0, sticky="ew", pady=5, padx=4)
            row.columnconfigure(2, weight=1)

            var = tk.BooleanVar(value=task["done"])
            ttk.Checkbutton(row, variable=var, command=self._on_toggle).grid(row=0, column=0, padx=(0, 10))

            label = ttk.Label(row, text=task["name"])
            if task["done"]:
                label.configure(foreground="#6b7280")
            label.grid(row=0, column=1, sticky="w", padx=(0, 10))

            note_var = tk.StringVar(value=task.get("notes", ""))
            ttk.Entry(row, textvariable=note_var).grid(row=0, column=2, sticky="ew")

            self._task_vars.append((var, label))
            self._note_vars.append(note_var)

        # Botones
        actions = tk.Frame(card, bg=CARD_BG)
        actions.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        self.finish_btn = ttk.Button(actions, text="Finalizar incorporación (CSV)",
                                     style="Primary.TButton", command=self.finish)
        self.finish_btn.pack(side="left")
        ttk.Button(actions, text="Reiniciar lista", command=self.reset).pack(side="left", padx=8)
        ttk.Button(actions, text="Guardar progreso", command=self._save_now).pack(side="right")
        ttk.Button(actions, text="Generar documento de material", command=self._generate_material).pack(side="right", padx=8)


        # Estado inicial de botones/progreso
        self._update_progress()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # -------- Helpers de UI / Scroll ----------
    def _on_mousewheel(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _on_mousewheel_linux(self, d):
        self.canvas.yview_scroll(d, "units")

    # -------- Progreso (robusto antes/después de crear checkboxes) ----------
    def _progress_counts(self):
        """Devuelve (done, total) de forma segura aunque _task_vars no exista aún."""
        if getattr(self, "_task_vars", None):
            done = sum(1 for v, _ in self._task_vars if v.get())
            total = len(self._task_vars)
        else:
            done = sum(1 for t in self.state_data["tasks"] if t.get("done"))
            total = len(self.state_data["tasks"])
        return done, total or 1

    def _progress_pct(self):
        done, total = self._progress_counts()
        return int(done / total * 100)

    def _progress_text(self):
        done, total = self._progress_counts()
        return f"Progreso: {done}/{total} · {self._progress_pct()}%"

    def _update_progress(self):
        self.progress_var.set(self._progress_pct())
        self.progress_label.configure(text=self._progress_text())
        self._update_finish_state()

    def _update_finish_state(self):
        all_done = all(v.get() for v, _ in self._task_vars) if self._task_vars else False
        has_name = bool(self.name_var.get().strip())
        self.finish_btn.state(["!disabled"] if (all_done and has_name) else ["disabled"])

    # -------- Persistencia ----------
    def _autosave(self):
        state = {
            "joinerName": self.name_var.get(),
            "technician": self.tech_var.get(),
            "requestNumber": self.req_var.get(),
            "date": self.date_var.get(),
            "tasks": [
                {"name": BASE_TASKS[i], "done": v.get(), "notes": self._note_vars[i].get()}
                for i, (v, _) in enumerate(self._task_vars)
            ],
        }
        save_state(state)

    def _save_now(self):
        self._autosave()
        messagebox.showinfo("Guardado", "Progreso guardado en state.json.")

    # -------- Acciones ----------
    def _on_toggle(self):
        for v, label in self._task_vars:
            label.configure(foreground="#6b7280" if v.get() else "#111827")
        self._update_progress()
        self._autosave()

    def reset(self):
        if not messagebox.askyesno("Reiniciar", "¿Seguro que quieres reiniciar la lista?"):
            return
        self.name_var.set("")
        self.tech_var.set("")
        self.req_var.set("")
        self.date_var.set(datetime.now().strftime("%Y-%m-%d"))
        for v, label in self._task_vars:
            v.set(False)
            label.configure(foreground="#111827")
        for nv in self._note_vars:
            nv.set("")
        self._update_progress()
        self._autosave()

    def finish(self):
        if not all(v.get() for v, _ in self._task_vars):
            messagebox.showwarning("Incompleto", "Debes marcar todas las tareas.")
            return
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Nombre requerido", "Indica el nombre.")
            return

        rows = [
            ("NOMBRE INCORPORACIÓN", name),
            ("FECHA", self.date_var.get().strip()),
            ("TÉCNICO", self.tech_var.get().strip()),
            ("NÚMERO DE PETICIÓN", self.req_var.get().strip()),
            ("PROGRESO", f"{self._progress_pct()}%"),
            ("REV", "01"),
            ("---", "---"),
        ]
        for i, (v, _) in enumerate(self._task_vars):
            rows.append((BASE_TASKS[i], "Hecho" if v.get() else "Pendiente"))

        safe = re.sub(r"[^\w\-\s.]", "_", name).strip().replace(" ", "_") or "INCORPORACION"
        default_name = f"{safe}.csv"
        path = filedialog.asksaveasfilename(
            title="Guardar checklist", defaultextension=".csv",
            initialfile=default_name, filetypes=[("CSV", "*.csv")]
        )
        if not path:
            return

        import csv
        with open(path, "w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerows(rows)

        self.reset()
        messagebox.showinfo("Guardado", f"Checklist guardado en:\n{path}\nSe reinició la lista.")

    def _on_close(self):
        self._autosave()
        self.destroy()

if __name__ == "__main__":
    App().mainloop()

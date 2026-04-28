import tkinter as tk
from tkinter import ttk, messagebox
import threading
import queue
import csv
import os
import sys
import time
import json
import subprocess
import importlib
import webbrowser
import shutil
import platform
import urllib.request
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import base64
import io
import math
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta


def _install_startup_package(package_name):
    """Install a package before the rest of the app imports it."""
    base_cmd = [sys.executable, "-m", "pip", "install", "--upgrade", package_name]
    try:
        result = subprocess.run(
            base_cmd,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except Exception as exc:
        raise RuntimeError(f"Could not start pip for {package_name}: {exc}") from exc

    if result.returncode == 0:
        return

    user_cmd = [sys.executable, "-m", "pip", "install", "--user", "--upgrade", package_name]
    result = subprocess.run(
        user_cmd,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip() or "pip returned an error."
        raise RuntimeError(f"Could not install {package_name}: {details[-2500:]}")


def _ensure_startup_import(import_name, package_name):
    """Import a required startup module; install it automatically if missing."""
    try:
        return importlib.import_module(import_name)
    except Exception:
        _install_startup_package(package_name)
        return importlib.import_module(import_name)


_ensure_startup_import("serial", "pyserial")
_ensure_startup_import("PIL", "Pillow")

import serial
import serial.tools.list_ports
from PIL import Image, ImageTk

gspread = None
Credentials = None
_HAS_GSPREAD = False
_GSPREAD_MISSING = []
_GSPREAD_IMPORT_ERRORS = []

SHEETS_REQUIRED_PACKAGES = ["gspread", "google-auth"]
MAP_REQUIRED_PACKAGES = ["tkintermapview"]

# ---------------------------------------------------------------------------
# Built-in Google Sheets configuration
# ---------------------------------------------------------------------------
# Fill these in if you want the SHEETS button to start immediately without
# asking for the Sheet URL and service-account JSON path.
#
# Example:
#   DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"
#   DEFAULT_SERVICE_ACCOUNT_JSON = r"C:\Users\laurm\Documents\cansat-service-account.json"
#
# Do not paste the private_key JSON contents directly into this Python file.
# Keep the JSON as a separate file and point DEFAULT_SERVICE_ACCOUNT_JSON to it.
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1KMH23124r0dfmFmYDiJ3dQ1vFEoOMmUcdY_b5E81EX0/edit?usp=sharing"
DEFAULT_SERVICE_ACCOUNT_JSON = r""

# Optional: environment variables override the saved config file.
SHEETS_URL_ENV = "CANSAT_GOOGLE_SHEET_URL"
SHEETS_CREDS_ENV = "CANSAT_SERVICE_ACCOUNT_JSON"


def check_sheets_dependencies():
    """Import Google Sheets dependencies using the Python running this script."""
    global gspread, Credentials, _HAS_GSPREAD
    global _GSPREAD_MISSING, _GSPREAD_IMPORT_ERRORS

    missing = []
    errors = []
    imported_gspread = None
    imported_credentials = None

    try:
        import gspread as imported_gspread
    except Exception as exc:
        missing.append("gspread")
        errors.append(f"gspread: {type(exc).__name__}: {exc}")

    try:
        from google.oauth2.service_account import Credentials as imported_credentials
    except Exception as exc:
        missing.append("google-auth")
        errors.append(f"google-auth: {type(exc).__name__}: {exc}")

    gspread = imported_gspread
    Credentials = imported_credentials
    _GSPREAD_MISSING = missing
    _GSPREAD_IMPORT_ERRORS = errors
    _HAS_GSPREAD = not missing
    return _HAS_GSPREAD, missing, errors


def sheets_dependency_message(missing=None, errors=None):
    missing = missing if missing is not None else _GSPREAD_MISSING
    errors = errors if errors is not None else _GSPREAD_IMPORT_ERRORS
    packages = " ".join(dict.fromkeys(missing or SHEETS_REQUIRED_PACKAGES))
    install_cmd = f'"{sys.executable}" -m pip install --upgrade {packages}'

    msg = (
        "Google Sheets dependencies could not be imported by the Python "
        "interpreter that is running this app.\n\n"
        f"Python in use:\n  {sys.executable}\n\n"
        "The app tried to install them automatically. If it still fails, "
        "run this exact command in PowerShell / CMD:\n\n"
        f"  {install_cmd}"
    )
    if errors:
        msg += "\n\nImport error details:\n  " + "\n  ".join(errors)
    return msg


def install_sheets_dependencies(missing=None, log_fn=None):
    """Install missing Google Sheets packages into the current Python."""
    log = log_fn or (lambda msg: None)
    packages = list(dict.fromkeys(missing or SHEETS_REQUIRED_PACKAGES))
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages

    log("Installing Google Sheets dependencies into current Python...")
    log("Command: " + " ".join(f'"{x}"' if " " in x else x for x in cmd))

    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    try:
        result = subprocess.run(
            cmd,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except Exception as exc:
        return False, f"Could not start pip: {type(exc).__name__}: {exc}"

    if result.stdout.strip():
        log(result.stdout.strip()[-1800:])
    if result.stderr.strip():
        log(result.stderr.strip()[-1800:])

    if result.returncode != 0:
        # A normal install can fail when Python is installed under Program Files.
        # Retry as a per-user install before giving up.
        user_cmd = [sys.executable, "-m", "pip", "install", "--user", "--upgrade"] + packages
        log("Normal install failed; retrying as current user...")
        log("Command: " + " ".join(f'"{x}"' if " " in x else x for x in user_cmd))
        result = subprocess.run(
            user_cmd,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if result.stdout.strip():
            log(result.stdout.strip()[-1800:])
        if result.stderr.strip():
            log(result.stderr.strip()[-1800:])
        if result.returncode != 0:
            details = result.stderr.strip() or result.stdout.strip() or "pip returned an error."
            return False, details[-2500:]

    return True, "Installed successfully."


def ensure_sheets_dependencies(parent=None, log_fn=None, auto_install=True):
    """Check Sheets imports and optionally install missing packages."""
    ok, missing, errors = check_sheets_dependencies()
    if ok:
        return True

    log = log_fn or (lambda msg: None)
    log("Google Sheets packages missing: " + ", ".join(missing))

    if not auto_install:
        return False

    if parent is not None:
        try:
            parent.config(cursor="watch")
            parent.update_idletasks()
        except Exception:
            pass

    installed, install_details = install_sheets_dependencies(missing, log_fn=log)

    if parent is not None:
        try:
            parent.config(cursor="")
            parent.update_idletasks()
        except Exception:
            pass

    if not installed:
        log("Automatic dependency install failed: " + install_details)
        return False

    ok, missing, errors = check_sheets_dependencies()
    if ok:
        log("Google Sheets dependencies installed and imported successfully.")
        return True

    log("Dependencies installed, but import still failed.")
    return False


def install_python_packages(packages, label, log_fn=None):
    """Install packages into the current Python, with a --user fallback."""
    log = log_fn or (lambda msg: None)
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages

    log(f"Installing {label} dependencies into current Python...")
    log("Command: " + " ".join(f'"{x}"' if " " in x else x for x in cmd))

    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    result = subprocess.run(
        cmd,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        startupinfo=startupinfo,
        creationflags=creationflags,
    )
    if result.stdout.strip():
        log(result.stdout.strip()[-1800:])
    if result.stderr.strip():
        log(result.stderr.strip()[-1800:])

    if result.returncode != 0:
        user_cmd = [sys.executable, "-m", "pip", "install", "--user", "--upgrade"] + packages
        log("Normal install failed; retrying as current user...")
        log("Command: " + " ".join(f'"{x}"' if " " in x else x for x in user_cmd))
        result = subprocess.run(
            user_cmd,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if result.stdout.strip():
            log(result.stdout.strip()[-1800:])
        if result.stderr.strip():
            log(result.stderr.strip()[-1800:])

    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip() or "pip returned an error."
        return False, details[-2500:]
    return True, "Installed successfully."


def ensure_map_dependencies(parent=None, log_fn=None, auto_install=True):
    """Install/import tkintermapview so the satellite map can use real tiles."""
    global tkintermapview, _HAS_MAPVIEW
    log = log_fn or (lambda msg: None)

    try:
        import tkintermapview as imported_mapview
        tkintermapview = imported_mapview
        _HAS_MAPVIEW = True
        return True
    except Exception as exc:
        _HAS_MAPVIEW = False
        log(f"Map package missing: tkintermapview ({type(exc).__name__}: {exc})")

    if not auto_install:
        return False

    if parent is not None:
        try:
            parent.config(cursor="watch")
            parent.update_idletasks()
        except Exception:
            pass

    try:
        installed, details = install_python_packages(
            MAP_REQUIRED_PACKAGES,
            "map",
            log_fn=log,
        )
    except Exception as exc:
        installed, details = False, f"Could not start pip: {type(exc).__name__}: {exc}"

    if parent is not None:
        try:
            parent.config(cursor="")
            parent.update_idletasks()
        except Exception:
            pass

    if not installed:
        log("Automatic map dependency install failed: " + details)
        return False

    try:
        import tkintermapview as imported_mapview
        tkintermapview = imported_mapview
        _HAS_MAPVIEW = True
        log("Map dependencies installed and imported successfully.")
        return True
    except Exception as exc:
        _HAS_MAPVIEW = False
        log(f"Map dependency installed, but import still failed: {type(exc).__name__}: {exc}")
        return False


check_sheets_dependencies()

try:
    import tkintermapview
    _HAS_MAPVIEW = True
except ImportError:
    tkintermapview = None
    _HAS_MAPVIEW = False

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False

try:
    from pyopengltk import OpenGLFrame
    from OpenGL.GL import (
        glClearColor, glEnable, glDisable, glClear, glBegin, glEnd,
        glVertex3f, glNormal3f, glColor3f, glColorMaterial,
        glLightfv, glMaterialfv, glMaterialf, glShadeModel,
        glMatrixMode, glLoadIdentity, glViewport,
        glRotatef, glScalef, glGenLists, glNewList, glEndList, glCallList,
        glPolygonMode, glPolygonOffset, glLineWidth,
        GL_DEPTH_TEST, GL_LIGHTING, GL_LIGHT0, GL_LIGHT1,
        GL_COLOR_MATERIAL, GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE,
        GL_SMOOTH, GL_POSITION, GL_DIFFUSE, GL_AMBIENT, GL_SPECULAR,
        GL_SHININESS, GL_TRIANGLES, GL_COMPILE,
        GL_COLOR_BUFFER_BIT, GL_DEPTH_BUFFER_BIT,
        GL_PROJECTION, GL_MODELVIEW,
        GL_POLYGON_OFFSET_FILL, GL_LINE, GL_FILL,
    )
    from OpenGL.GLU import gluPerspective, gluLookAt
    _HAS_GL = True
except ImportError:
    _HAS_GL = False


def load_3mf(filepath):
    ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02"
    vertices, triangles = [], []
    with zipfile.ZipFile(filepath, "r") as zf:
        model_files = [n for n in zf.namelist() if n.endswith(".model")]
        if not model_files:
            raise ValueError("No .model file found in .3mf archive")
        with zf.open(model_files[0]) as f:
            root = ET.parse(f).getroot()

    for mesh in root.iter(f"{{{ns}}}mesh"):
        base = len(vertices)
        verts_el = mesh.find(f"{{{ns}}}vertices")
        if verts_el is not None:
            for v in verts_el.findall(f"{{{ns}}}vertex"):
                vertices.append((
                    float(v.get("x", 0)),
                    float(v.get("y", 0)),
                    float(v.get("z", 0)),
                ))
        tris_el = mesh.find(f"{{{ns}}}triangles")
        if tris_el is not None:
            for t in tris_el.findall(f"{{{ns}}}triangle"):
                triangles.append((
                    base + int(t.get("v1")),
                    base + int(t.get("v2")),
                    base + int(t.get("v3")),
                ))
    return vertices, triangles


if _HAS_GL:
    class CanSatGLFrame(OpenGLFrame):
        def __init__(self, parent, app, model_path, **kwargs):
            super().__init__(parent, **kwargs)
            self.app = app
            self.model_path = model_path
            self._display_list = None
            self._model_scale = 1.0
            self.animate = 33

            self.imu_driven = True
            self._mouse_rx = 0.0
            self._mouse_ry = 0.0
            self._drag_last = None

            self.bind("<ButtonPress-1>", self._on_mouse_press)
            self.bind("<B1-Motion>", self._on_mouse_drag)

        def initgl(self):
            glClearColor(0.01, 0.01, 0.04, 1.0)
            glEnable(GL_DEPTH_TEST)
            glEnable(GL_LIGHTING)
            glEnable(GL_LIGHT0)
            glEnable(GL_LIGHT1)
            glEnable(GL_COLOR_MATERIAL)
            glColorMaterial(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE)
            glShadeModel(GL_SMOOTH)

            glLightfv(GL_LIGHT0, GL_POSITION, [3.0, 5.0, 4.0, 0.0])
            glLightfv(GL_LIGHT0, GL_DIFFUSE, [0.0, 0.9, 0.85, 1.0])
            glLightfv(GL_LIGHT0, GL_AMBIENT, [0.04, 0.04, 0.08, 1.0])
            glLightfv(GL_LIGHT0, GL_SPECULAR, [0.0, 1.0, 0.9, 1.0])

            glLightfv(GL_LIGHT1, GL_POSITION, [-3.0, -2.0, -3.0, 0.0])
            glLightfv(GL_LIGHT1, GL_DIFFUSE, [0.6, 0.0, 0.55, 1.0])
            glLightfv(GL_LIGHT1, GL_AMBIENT, [0.0, 0.0, 0.0, 1.0])

            glMaterialfv(GL_FRONT_AND_BACK, GL_SPECULAR, [0.0, 1.0, 0.9, 1.0])
            glMaterialf(GL_FRONT_AND_BACK, GL_SHININESS, 90.0)

            self._build_display_list()

        def redraw(self):
            w = max(self.winfo_width(), 1)
            h = max(self.winfo_height(), 1)
            glViewport(0, 0, w, h)
            glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)

            glMatrixMode(GL_PROJECTION)
            glLoadIdentity()
            gluPerspective(45.0, w / h, 0.1, 100.0)

            glMatrixMode(GL_MODELVIEW)
            glLoadIdentity()
            gluLookAt(0, 0, 4, 0, 0, 0, 0, 1, 0)

            s = self._model_scale
            glScalef(s, s, s)

            if self.imu_driven:
                glRotatef(self.app.roll, 0, 1, 0)
                glRotatef(-self.app.pitch, 1, 0, 0)
            else:
                glRotatef(self._mouse_ry, 0, 1, 0)
                glRotatef(self._mouse_rx, 1, 0, 0)
            glRotatef(-90, 1, 0, 0)

            if self._display_list:
                glEnable(GL_POLYGON_OFFSET_FILL)
                glPolygonOffset(2.0, 2.0)
                glPolygonMode(GL_FRONT_AND_BACK, GL_FILL)
                glColor3f(0.03, 0.06, 0.14)
                glCallList(self._display_list)
                glDisable(GL_POLYGON_OFFSET_FILL)

                glPolygonMode(GL_FRONT_AND_BACK, GL_LINE)
                glDisable(GL_LIGHTING)
                glColor3f(0.0, 0.95, 0.88)
                glLineWidth(0.8)
                glCallList(self._display_list)
                glPolygonMode(GL_FRONT_AND_BACK, GL_FILL)
                glEnable(GL_LIGHTING)
                glLineWidth(1.0)

        def _on_mouse_press(self, event):
            self._drag_last = (event.x, event.y)

        def _on_mouse_drag(self, event):
            if self.imu_driven or self._drag_last is None:
                return
            dx = event.x - self._drag_last[0]
            dy = event.y - self._drag_last[1]
            self._mouse_ry += dx * 0.5
            self._mouse_rx += dy * 0.5
            self._drag_last = (event.x, event.y)

        def _build_display_list(self):
            try:
                verts, tris = load_3mf(self.model_path)
            except Exception as e:
                print(f"[3D] Failed to load model: {e}")
                return

            if not verts:
                return

            xs = [v[0] for v in verts]
            ys = [v[1] for v in verts]
            zs = [v[2] for v in verts]
            cx = (max(xs) + min(xs)) / 2.0
            cy = (max(ys) + min(ys)) / 2.0
            cz = (max(zs) + min(zs)) / 2.0
            extent = max(max(xs) - min(xs), max(ys) - min(ys), max(zs) - min(zs), 1e-9)
            self._model_scale = 1.8 / extent

            normals = []
            for t in tris:
                v0, v1, v2 = verts[t[0]], verts[t[1]], verts[t[2]]
                e1 = (v1[0] - v0[0], v1[1] - v0[1], v1[2] - v0[2])
                e2 = (v2[0] - v0[0], v2[1] - v0[1], v2[2] - v0[2])
                nx = e1[1] * e2[2] - e1[2] * e2[1]
                ny = e1[2] * e2[0] - e1[0] * e2[2]
                nz = e1[0] * e2[1] - e1[1] * e2[0]
                ln = math.sqrt(nx * nx + ny * ny + nz * nz) or 1e-9
                normals.append((nx / ln, ny / ln, nz / ln))

            self._display_list = glGenLists(1)
            glNewList(self._display_list, GL_COMPILE)
            glBegin(GL_TRIANGLES)
            for i, t in enumerate(tris):
                if i < len(normals):
                    glNormal3f(*normals[i])
                for vi in t:
                    v = verts[vi]
                    glVertex3f(v[0] - cx, v[1] - cy, v[2] - cz)
            glEnd()
            glEndList()


class TelemetryVizWindow:
    _BG = "#020208"
    _CYAN = "#00FFE5"
    _MAGENTA = "#FF0090"
    _DIM_CYAN = "#004D47"

    _CHART_COLORS = {
        "temp": "#FF0090",
        "pressure": "#00FFE5",
    }

    def __init__(self, app):
        self.app = app
        self.win = tk.Toplevel(app.root)
        self.win.title("CanSat — Telemetry Visualization")
        self.win.geometry("1300x750")
        self.win.configure(bg=self._BG)
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        self.time_window_sec = tk.IntVar(value=60)
        self._att_var = tk.StringVar(value="ROLL  +0.0°   PITCH  +0.0°")

        self._build_ui()
        self._schedule_chart_update()

    def _on_close(self):
        self.win.withdraw()

    def _build_ui(self):
        left = ttk.Frame(self.win, padding=(10, 10, 5, 10))
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(self.win, padding=(5, 10, 10, 10))
        right.pack(side="right", fill="both")

        self._build_charts(left)
        self._build_3d(right)

    def _build_charts(self, parent):
        ctrl = ttk.Frame(parent)
        ctrl.pack(fill="x", pady=(0, 6))

        ttk.Label(ctrl, text="Time window:", style="Header.TLabel").pack(side="left")

        self._slider_label = ttk.Label(ctrl, text="60 s", style="Value.TLabel", width=6)
        self._slider_label.pack(side="right")

        ttk.Scale(
            ctrl,
            from_=10, to=300,
            variable=self.time_window_sec,
            orient="horizontal",
            command=self._on_slider,
        ).pack(side="left", fill="x", expand=True, padx=8)

        if not _HAS_MPL:
            ttk.Label(
                parent,
                text="matplotlib not installed — charts unavailable",
                style="Header.TLabel",
            ).pack(expand=True)
            return

        fig = Figure(figsize=(7, 6), facecolor=self._BG)
        fig.subplots_adjust(hspace=0.42, left=0.11, right=0.97, top=0.93, bottom=0.08)

        self.ax_temp = fig.add_subplot(2, 1, 1)
        self.ax_press = fig.add_subplot(2, 1, 2)

        self._style_axes()

        self._canvas = FigureCanvasTkAgg(fig, master=parent)
        self._canvas.get_tk_widget().pack(fill="both", expand=True)

    def _style_axes(self):
        configs = [
            (self.ax_temp, "TEMPERATURE", self._CHART_COLORS["temp"], "°C"),
            (self.ax_press, "PRESSURE", self._CHART_COLORS["pressure"], "hPa"),
        ]
        for ax, title, color, ylabel in configs:
            ax.set_facecolor("#00020F")
            ax.set_title(title, color=color, fontsize=10, pad=6,
                         fontfamily="monospace", fontweight="bold")
            ax.set_ylabel(ylabel, color=color, fontsize=9, fontfamily="monospace")
            ax.tick_params(colors=self._DIM_CYAN, labelsize=8)
            for side in ("bottom", "left"):
                ax.spines[side].set_color(self._DIM_CYAN)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(True, color="#001A18", linewidth=0.8)

    def _build_3d(self, parent):
        lf = ttk.LabelFrame(parent, text="Satellite Position", padding=5)
        lf.pack(fill="both", expand=True)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        model_path = os.path.join(base_dir, "Purk V2.3mf")

        bar = tk.Frame(lf, bg=self._BG, pady=4)
        bar.pack(side="bottom", fill="x")

        self._mode_btn = tk.Button(
            bar,
            text="[ IMU ]",
            bg=self._CYAN,
            fg=self._BG,
            activebackground="#33FFF0",
            activeforeground=self._BG,
            font=("Consolas", 11, "bold"),
            relief="flat",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=self._toggle_mode,
        )
        self._mode_btn.pack(side="left", padx=(4, 10))

        tk.Label(
            bar,
            textvariable=self._att_var,
            bg=self._BG,
            fg=self._CYAN,
            font=("Consolas", 10, "bold"),
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

        if not _HAS_GL:
            ttk.Label(
                lf,
                text="Install pyopengltk + PyOpenGL\nfor 3-D visualization",
                style="Header.TLabel",
                justify="center",
            ).pack(expand=True)
            return

        if not os.path.exists(model_path):
            ttk.Label(
                lf,
                text=f"3-D model not found:\n{model_path}",
                style="Header.TLabel",
                justify="center",
            ).pack(expand=True)
            return

        self._gl = CanSatGLFrame(lf, self.app, model_path, width=480, height=540)
        self._gl.pack(fill="both", expand=True)

    def _toggle_mode(self):
        if not hasattr(self, "_gl"):
            return

        gl = self._gl
        gl.imu_driven = not gl.imu_driven
        if gl.imu_driven:
            self._mode_btn.config(
                text="[ IMU ]",
                bg=self._CYAN, fg=self._BG,
                activebackground="#33FFF0",
            )
        else:
            gl._mouse_rx = -self.app.pitch
            gl._mouse_ry = self.app.roll
            self._mode_btn.config(
                text="[ MOUSE ]",
                bg=self._MAGENTA, fg=self._BG,
                activebackground="#FF33AA",
            )

    def _on_slider(self, val):
        self._slider_label.config(text=f"{int(float(val))} s")

    def _schedule_chart_update(self):
        self._update_charts()
        self.win.after(500, self._schedule_chart_update)

    def _update_charts(self):
        if not self.win.winfo_exists():
            return

        if hasattr(self, "_gl") and not self._gl.imu_driven:
            self._att_var.set(
                f"AZ  {self._gl._mouse_ry:+.1f}°   EL  {self._gl._mouse_rx:+.1f}°  [mouse]"
            )
        else:
            self._att_var.set(
                f"ROLL  {self.app.roll:+.1f}°   PITCH  {self.app.pitch:+.1f}°"
            )

        if not _HAS_MPL:
            return

        history = self.app.telemetry_history
        if not history:
            return

        secs = self.time_window_sec.get()
        cutoff = datetime.now() - timedelta(seconds=secs)
        visible = [d for d in history if d["time"] >= cutoff]
        if not visible:
            return

        t0 = visible[0]["time"]
        x = [(d["time"] - t0).total_seconds() for d in visible]
        temps = [d["temp"] for d in visible]
        pressures = [d["pressure"] for d in visible]

        for ax, data, color, title, ylabel, xlabel in (
            (self.ax_temp, temps, self._CHART_COLORS["temp"], "Temperature", "°C", ""),
            (self.ax_press, pressures, self._CHART_COLORS["pressure"], "Pressure", "hPa", "seconds"),
        ):
            ax.clear()
            ax.set_facecolor("#00020F")
            ax.set_title(title, color=color, fontsize=10, pad=6,
                         fontfamily="monospace", fontweight="bold")
            ax.set_ylabel(ylabel, color=color, fontsize=9, fontfamily="monospace")
            if xlabel:
                ax.set_xlabel(xlabel, color=self._DIM_CYAN, fontsize=8,
                              fontfamily="monospace")
            ax.tick_params(colors=self._DIM_CYAN, labelsize=8)
            for side in ("bottom", "left"):
                ax.spines[side].set_color(self._DIM_CYAN)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(True, color="#001A18", linewidth=0.8)

            ax.plot(x, data, color=color, linewidth=2, zorder=3)
            ax.fill_between(x, data, alpha=0.15, color=color, zorder=2)

            if data:
                ax.annotate(
                    f"{data[-1]:.2f}",
                    xy=(x[-1], data[-1]),
                    xytext=(6, 4),
                    textcoords="offset points",
                    color=color,
                    fontsize=9,
                    fontweight="bold",
                    fontfamily="monospace",
                )

        self._canvas.draw_idle()


class AttitudeIndicator(tk.Canvas):
    def __init__(self, parent, app, width=420, height=420, **kwargs):
        super().__init__(
            parent,
            width=width,
            height=height,
            bg="#111111",
            highlightthickness=0,
            **kwargs
        )
        self.app = app
        self.w = width
        self.h = height
        self.cx = width / 2
        self.cy = height / 2
        self.radius = min(width, height) / 2 - 12

        self.sky_color = "#3A6DB4"
        self.ground_color = "#8B5A2B"
        self.hud_color = "#F0F0F0"
        self.accent_color = "#F8E45C"
        self.vv_color = "#7CFF7C"

    def _rot(self, x, y, deg):
        a = math.radians(deg)
        xr = x * math.cos(a) - y * math.sin(a)
        yr = x * math.sin(a) + y * math.cos(a)
        return xr, yr

    def _draw_bank_marks(self, cx, cy, r):
        bank_marks = [-60, -45, -30, -20, -10, 10, 20, 30, 45, 60]
        for ang in bank_marks:
            inner = r - 12
            outer = r - (26 if abs(ang) in (30, 60) else 19)
            x1 = cx + inner * math.sin(math.radians(ang))
            y1 = cy - inner * math.cos(math.radians(ang))
            x2 = cx + outer * math.sin(math.radians(ang))
            y2 = cy - outer * math.cos(math.radians(ang))
            self.create_line(x1, y1, x2, y2, fill=self.hud_color, width=2)

        pointer = [
            (cx, cy - r + 10),
            (cx - 8, cy - r + 24),
            (cx + 8, cy - r + 24),
        ]
        self.create_polygon(pointer, fill=self.accent_color, outline="")

    def _draw_fixed_aircraft_symbol(self, cx, cy):
        self.create_line(cx - 50, cy, cx - 14, cy, fill=self.hud_color, width=3)
        self.create_line(cx + 14, cy, cx + 50, cy, fill=self.hud_color, width=3)
        self.create_line(cx, cy - 10, cx, cy + 10, fill=self.hud_color, width=2)
        self.create_rectangle(cx - 4, cy - 4, cx + 4, cy + 4, outline=self.hud_color)

    def _draw_velocity_vector(self, cx, cy):
        vx = cx + self.app.vx
        vy = cy + self.app.vy
        self.create_oval(vx - 8, vy - 8, vx + 8, vy + 8, outline=self.vv_color, width=2)
        self.create_line(vx - 16, vy, vx - 8, vy, fill=self.vv_color, width=2)
        self.create_line(vx + 8, vy, vx + 16, vy, fill=self.vv_color, width=2)
        self.create_line(vx, vy - 16, vx, vy - 8, fill=self.vv_color, width=2)

    def draw_indicator(self, roll_deg, pitch_deg):
        self.delete("all")

        cx, cy, r = self.cx, self.cy, self.radius
        pitch_scale = 3.0
        pitch_px = pitch_deg * pitch_scale
        scale = 6 * r

        sky_rect = [
            (-scale, -scale - pitch_px),
            (scale, -scale - pitch_px),
            (scale, -pitch_px),
            (-scale, -pitch_px),
        ]

        ground_rect = [
            (-scale, -pitch_px),
            (scale, -pitch_px),
            (scale, scale - pitch_px),
            (-scale, scale - pitch_px),
        ]

        def transform(points):
            out = []
            for x, y in points:
                xr, yr = self._rot(x, y, -roll_deg)
                out.extend([cx + xr, cy + yr])
            return out

        self.create_polygon(transform(sky_rect), fill=self.sky_color, outline="", smooth=False)
        self.create_polygon(transform(ground_rect), fill=self.ground_color, outline="", smooth=False)

        x1, y1 = self._rot(-140, -pitch_px, -roll_deg)
        x2, y2 = self._rot(140, -pitch_px, -roll_deg)
        self.create_line(cx + x1, cy + y1, cx + x2, cy + y2, fill=self.accent_color, width=3)

        for mark in range(-30, 31, 5):
            if mark == 0:
                continue

            y = -mark * pitch_scale - pitch_px
            half = 38 if mark % 10 == 0 else 22

            lx1, ly1 = self._rot(-half, y, -roll_deg)
            lx2, ly2 = self._rot(half, y, -roll_deg)
            self.create_line(cx + lx1, cy + ly1, cx + lx2, cy + ly2, fill=self.hud_color, width=2)

            if mark % 10 == 0:
                txl, tyl = self._rot(-half - 18, y, -roll_deg)
                txr, tyr = self._rot(half + 18, y, -roll_deg)
                label = str(abs(mark))
                self.create_text(cx + txl, cy + tyl, text=label, fill=self.hud_color, font=("Segoe UI", 10))
                self.create_text(cx + txr, cy + tyr, text=label, fill=self.hud_color, font=("Segoe UI", 10))

        self._draw_bank_marks(cx, cy, r)
        self._draw_fixed_aircraft_symbol(cx, cy)
        self._draw_velocity_vector(cx, cy)

        self.create_oval(cx - r, cy - r, cx + r, cy + r, outline="#DDDDDD", width=3)

        self.create_text(
            cx,
            cy + r - 18,
            text=f"ROLL {roll_deg:+05.1f}°   PITCH {pitch_deg:+05.1f}°",
            fill="#DDDDDD",
            font=("Segoe UI", 10, "bold")
        )


# ---------------------------------------------------------------------------
# Google Sheets live uploader
# ---------------------------------------------------------------------------

class GoogleSheetsUploader:
    """
    Streams new CSV rows to a Google Spreadsheet in a background thread.

    Two worksheets are maintained:
      • "Telemetry"  – data from telemetry.csv
      • "GPS"        – data from gps.csv

    After the header row is written, three charts are added once to the
    Telemetry sheet:
      1. Temperature vs time
      2. Pressure vs time
      3. Altitude vs time

    Usage
    -----
    uploader = GoogleSheetsUploader(
        sheet_url   = "https://docs.google.com/spreadsheets/d/…",
        credentials = "/path/to/service_account.json",
        log_fn      = some_callable_that_accepts_a_string,
    )
    uploader.start()

    # Call whenever a new telemetry / GPS row has been appended to the CSV:
    uploader.notify()

    uploader.stop()
    """

    _TELEMETRY_HEADERS = [
        "row_id", "mission_id", "source",
        "seq", "timestamp", "temp_C", "pressure_hPa", "alt_m",
        "ax", "ay", "az", "gx", "gy", "gz",
        "peak_alt_m", "descent_count", "deployed", "deploy_reason",
    ]
    _GPS_HEADERS = [
        "row_id", "mission_id", "source",
        "seq", "timestamp", "latitude", "longitude",
        "gnss_alt_m", "satellites", "fix_quality",
    ]

    # Google Sheets has per-minute write quotas. Keep uploads batched so
    # live telemetry does not call the API once per sensor packet.
    _MIN_UPLOAD_INTERVAL_SEC = 15.0
    _MAX_ROWS_PER_APPEND = 500
    _INITIAL_BACKOFF_SEC = 10.0
    _MAX_BACKOFF_SEC = 120.0

    def __init__(self, sheet_url: str, credentials: str, telemetry_csv: str,
                 gps_csv: str, telemetry_test_csv: str = None,
                 gps_test_csv: str = None, log_fn=None):
        self.sheet_url = sheet_url
        self.credentials_path = credentials
        self.telemetry_csv = telemetry_csv
        self.gps_csv = gps_csv
        self.telemetry_test_csv = telemetry_test_csv
        self.gps_test_csv = gps_test_csv
        self.log = log_fn or (lambda msg: print(f"[Sheets] {msg}"))

        self._stop_event = threading.Event()
        self._notify_event = threading.Event()
        self._thread = None

        # Track how many data rows (excluding header) we have already uploaded
        self._tel_sent = 0
        self._gps_sent = 0
        self._tel_test_sent = 0
        self._gps_test_sent = 0

        # gspread worksheet handles, set once connected
        self._ws_tel = None
        self._ws_gps = None
        self._ws_tel_test = None
        self._ws_gps_test = None
        self._spreadsheet = None

        # Upload throttling / 429 retry state.
        self._last_flush_monotonic = 0.0
        self._backoff_until = 0.0
        self._backoff_seconds = self._INITIAL_BACKOFF_SEC
        self._last_wait_log = 0.0

        # Local upload state is used instead of using the number of rows that
        # already exist in the Google Sheet. That older approach can make a
        # fresh local GPS CSV look "already uploaded" when the GPS worksheet
        # contains rows from a previous run.
        self._state_path = self._default_state_path()

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def start(self):
        if not ensure_sheets_dependencies(log_fn=self.log, auto_install=True):
            self.log(sheets_dependency_message())
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True,
                                        name="SheetsUploader")
        self._thread.start()
        self.log("Sheets uploader thread started.")

    def stop(self):
        self._stop_event.set()
        self._notify_event.set()   # unblock the wait
        if self._thread:
            self._thread.join(timeout=5)
        self.log("Sheets uploader stopped.")

    def notify(self):
        """Call this whenever new rows have been written to either CSV."""
        self._notify_event.set()

    # ------------------------------------------------------------------
    # Background thread
    # ------------------------------------------------------------------

    def _run(self):
        try:
            self._connect()
        except Exception as exc:
            self.log(f"Sheets connect failed: {exc}")
            return

        # Allow the first upload immediately, then throttle later uploads.
        self._last_flush_monotonic = time.monotonic() - self._MIN_UPLOAD_INTERVAL_SEC

        while True:
            # Wake on a new row, or periodically in case a notify was missed.
            self._notify_event.wait(timeout=self._MIN_UPLOAD_INTERVAL_SEC)
            self._notify_event.clear()

            stopping = self._stop_event.is_set()
            now = time.monotonic()

            # If Google returned 429 recently, wait before trying again.
            if not stopping and now < self._backoff_until:
                self._log_wait_once(
                    f"Sheets quota backoff: retrying in {self._backoff_until - now:.0f}s."
                )
                continue

            # Batch rows for a few seconds instead of writing once per packet.
            wait_remaining = self._MIN_UPLOAD_INTERVAL_SEC - (now - self._last_flush_monotonic)
            if not stopping and wait_remaining > 0:
                self._log_wait_once(
                    f"Sheets batching rows; next upload in {wait_remaining:.0f}s."
                )
                continue

            # Flush GPS first so the GPS worksheet cannot be starved by a large
            # telemetry backlog or a telemetry-side quota failure. Each worksheet
            # is also handled independently, so one failure does not block the other.
            gps_uploaded, tel_uploaded = 0, 0
            had_quota_error = False

            try:
                gps_uploaded = self._flush_gps()
            except Exception as exc:
                if self._is_quota_error(exc):
                    had_quota_error = True
                    self.log("Sheets GPS upload hit quota [429]. GPS rows were not marked as uploaded.")
                else:
                    self.log(f"Sheets GPS upload error: {exc}")

            try:
                tel_uploaded = self._flush_telemetry()
            except Exception as exc:
                if self._is_quota_error(exc):
                    had_quota_error = True
                    self.log("Sheets telemetry upload hit quota [429]. Telemetry rows were not marked as uploaded.")
                else:
                    self.log(f"Sheets telemetry upload error: {exc}")

            self._last_flush_monotonic = time.monotonic()
            if gps_uploaded or tel_uploaded:
                self.log(
                    f"Sheets batch complete: telemetry={tel_uploaded}, gps={gps_uploaded}."
                )

            if had_quota_error:
                delay = self._backoff_seconds
                self._backoff_until = time.monotonic() + delay
                self._backoff_seconds = min(
                    self._backoff_seconds * 2, self._MAX_BACKOFF_SEC
                )
                self.log(
                    "Sheets quota backoff enabled; "
                    f"retrying in {delay:.0f}s."
                )
            else:
                self._backoff_seconds = self._INITIAL_BACKOFF_SEC

            if stopping:
                break

    def _log_wait_once(self, message):
        now = time.monotonic()
        if now - self._last_wait_log >= 10.0:
            self.log(message)
            self._last_wait_log = now

    def _is_quota_error(self, exc):
        text = str(exc).lower()
        return "429" in text or "quota exceeded" in text or "too many requests" in text

    def _default_state_path(self):
        """Return the local JSON file used to remember uploaded CSV rows."""
        try:
            csv_dir = os.path.dirname(os.path.abspath(self.telemetry_csv))
            output_root = os.path.dirname(csv_dir)
            secrets_dir = os.path.join(output_root, "Secrets")
        except Exception:
            secrets_dir = os.path.join(os.getcwd(), "CanSat_GroundStation_Data", "Secrets")
        return os.path.join(secrets_dir, "sheets_upload_state.json")

    def _load_upload_state(self):
        try:
            if os.path.isfile(self._state_path):
                with open(self._state_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as exc:
            self.log(f"Sheets upload state ignored: {exc}")
        return {}

    def _save_upload_state(self):
        try:
            os.makedirs(os.path.dirname(self._state_path), exist_ok=True)
            data = self._load_upload_state()
            sheet_id = self._spreadsheet.id if self._spreadsheet else "unknown_sheet"
            data.setdefault(sheet_id, {})
            data[sheet_id]["telemetry"] = self._state_entry(self.telemetry_csv, self._tel_sent)
            data[sheet_id]["gps"] = self._state_entry(self.gps_csv, self._gps_sent)
            if self.telemetry_test_csv:
                data[sheet_id]["telemetry_test"] = self._state_entry(self.telemetry_test_csv, self._tel_test_sent)
            if self.gps_test_csv:
                data[sheet_id]["gps_test"] = self._state_entry(self.gps_test_csv, self._gps_test_sent)
            tmp = self._state_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            os.replace(tmp, self._state_path)
        except Exception as exc:
            self.log(f"Sheets upload state save failed: {exc}")

    def _state_entry(self, csv_path, sent_rows):
        path = os.path.abspath(csv_path)
        try:
            stat = os.stat(path)
            mtime = stat.st_mtime
            size = stat.st_size
        except OSError:
            mtime = 0.0
            size = 0
        return {
            "csv_path": path,
            "sent_rows": int(max(0, sent_rows)),
            "csv_mtime": mtime,
            "csv_size": size,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }

    def _count_csv_data_rows(self, csv_path):
        if not os.path.exists(csv_path):
            return 0
        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                return max(0, sum(1 for _ in csv.reader(f)) - 1)
        except Exception as exc:
            self.log(f"Could not count local CSV rows for {os.path.basename(csv_path)}: {exc}")
            return 0

    def _counter_from_state(self, state, key, csv_path, local_rows):
        if not csv_path:
            return 0
        sheet_id = self._spreadsheet.id if self._spreadsheet else "unknown_sheet"
        entry = state.get(sheet_id, {}).get(key, {})
        saved_path = entry.get("csv_path")
        try:
            saved_count = int(entry.get("sent_rows", 0))
        except Exception:
            saved_count = 0

        # Reset if there is no usable state, if the CSV path changed, or if the
        # CSV now has fewer rows than the saved counter. That last case happens
        # when a fresh mission CSV is created while the Google Sheet still has
        # rows from an older run.
        if saved_path != os.path.abspath(csv_path):
            return 0
        if local_rows < saved_count:
            return 0
        return max(0, min(saved_count, local_rows))

    def _init_upload_counters(self):
        state = self._load_upload_state()
        tel_local = self._count_csv_data_rows(self.telemetry_csv)
        gps_local = self._count_csv_data_rows(self.gps_csv)
        tel_test_local = self._count_csv_data_rows(self.telemetry_test_csv) if self.telemetry_test_csv else 0
        gps_test_local = self._count_csv_data_rows(self.gps_test_csv) if self.gps_test_csv else 0
        self._tel_sent = self._counter_from_state(state, "telemetry", self.telemetry_csv, tel_local)
        self._gps_sent = self._counter_from_state(state, "gps", self.gps_csv, gps_local)
        self._tel_test_sent = self._counter_from_state(state, "telemetry_test", self.telemetry_test_csv, tel_test_local) if self.telemetry_test_csv else 0
        self._gps_test_sent = self._counter_from_state(state, "gps_test", self.gps_test_csv, gps_test_local) if self.gps_test_csv else 0

        try:
            tel_sheet = max(0, len(self._ws_tel.get_all_values()) - 1)
        except Exception:
            tel_sheet = -1
        try:
            gps_sheet = max(0, len(self._ws_gps.get_all_values()) - 1)
        except Exception:
            gps_sheet = -1

        self.log(
            "Sheets counters initialised from local state "
            f"(local telemetry={tel_local}, local gps={gps_local}, "
            f"local telemetry_TEST={tel_test_local}, local gps_TEST={gps_test_local}, "
            f"sheet telemetry={tel_sheet}, sheet gps={gps_sheet})."
        )

    def _connect(self):
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_file(
            self.credentials_path, scopes=scopes
        )
        gc = gspread.authorize(creds)

        # Open by URL first; fall back to extracting the spreadsheet key.
        try:
            self._spreadsheet = gc.open_by_url(self.sheet_url)
        except Exception:
            sheet_id = self.sheet_url.split("/d/")[1].split("/")[0]
            self._spreadsheet = gc.open_by_key(sheet_id)

        self._ws_tel = self._get_or_create_worksheet("Telemetry")
        self._ws_gps = self._get_or_create_worksheet("GPS")
        self._ws_tel_test = self._get_or_create_worksheet("Telemetry_TEST")
        self._ws_gps_test = self._get_or_create_worksheet("GPS_TEST")

        # Write headers if the sheet is empty
        self._ensure_header(self._ws_tel, self._TELEMETRY_HEADERS)
        self._ensure_header(self._ws_gps, self._GPS_HEADERS)
        self._ensure_header(self._ws_tel_test, self._TELEMETRY_HEADERS)
        self._ensure_header(self._ws_gps_test, self._GPS_HEADERS)

        self._init_upload_counters()

        self.log(f"Connected to Google Sheet. "
                 f"Local upload counters: telemetry={self._tel_sent}, "
                 f"gps={self._gps_sent}.")

        # Add charts only when the Telemetry worksheet has no uploaded data yet.
        # This avoids duplicate chart creation and saves Sheets write quota on restarts.
        tel_sheet_rows = max(0, len(self._ws_tel.get_all_values()) - 1)
        if tel_sheet_rows == 0:
            try:
                self._add_telemetry_charts()
            except Exception as exc:
                self.log(f"Chart creation skipped: {exc}")

    def _get_or_create_worksheet(self, title: str):
        try:
            return self._spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            return self._spreadsheet.add_worksheet(title=title,
                                                    rows=10000, cols=20)

    def _ensure_header(self, ws, headers):
        existing = ws.row_values(1)
        if not existing:
            ws.append_row(headers)
        elif existing[:len(headers)] != headers:
            # Keep existing data, but update the visible header row when the app
            # gains new columns such as row_id / mission_id. This is a single
            # write on connect, not a per-packet write.
            ws.update("1:1", [headers])

    # ------------------------------------------------------------------
    # CSV → Sheets flushing
    # ------------------------------------------------------------------

    def _flush_telemetry(self):
        uploaded = self._flush_csv(self.telemetry_csv, self._ws_tel,
                                   None, "_tel_sent", self._TELEMETRY_HEADERS)
        if self.telemetry_test_csv:
            uploaded += self._flush_csv(self.telemetry_test_csv, self._ws_tel_test,
                                        None, "_tel_test_sent", self._TELEMETRY_HEADERS)
        return uploaded

    def _flush_gps(self):
        uploaded = self._flush_csv(self.gps_csv, self._ws_gps,
                                   None, "_gps_sent", self._GPS_HEADERS)
        if self.gps_test_csv:
            uploaded += self._flush_csv(self.gps_test_csv, self._ws_gps_test,
                                        None, "_gps_test_sent", self._GPS_HEADERS)
        return uploaded

    def _flush_csv(self, csv_path: str, ws, test_ws, sent_attr: str, expected_headers):
        if not os.path.exists(csv_path):
            return 0

        with open(csv_path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)

        # rows[0] is the header; data starts at rows[1]
        if len(rows) < 2:
            return 0

        already_sent = getattr(self, sent_attr)
        new_rows = rows[1 + already_sent:]   # skip header + already-uploaded

        if not new_rows:
            return 0

        uploaded = 0
        write_requests = 0
        source_index = expected_headers.index("source") if "source" in expected_headers else -1

        for start in range(0, len(new_rows), self._MAX_ROWS_PER_APPEND):
            chunk = new_rows[start:start + self._MAX_ROWS_PER_APPEND]
            test_rows = [r for r in chunk if source_index >= 0 and len(r) > source_index and str(r[source_index]).upper() == "TEST"]

            main_start_row = self._next_append_row(ws)
            ws.append_rows(chunk, value_input_option="USER_ENTERED")
            write_requests += 1

            if test_rows and test_ws is not None:
                test_start_row = self._next_append_row(test_ws)
                test_ws.append_rows(test_rows, value_input_option="USER_ENTERED")
                write_requests += 1
                # Best-effort coloring for the split TEST worksheet too.
                self._color_sheet_rows(test_ws, test_start_row, test_start_row + len(test_rows) - 1, len(expected_headers), "test")

            # Best-effort coloring for TEST rows in the main mixed worksheet.
            if test_rows:
                for offset, row in enumerate(chunk):
                    if len(row) > source_index and str(row[source_index]).upper() == "TEST":
                        r = main_start_row + offset
                        self._color_sheet_rows(ws, r, r, len(expected_headers), "test")

            uploaded += len(chunk)
            setattr(self, sent_attr, already_sent + uploaded)
            self._save_upload_state()

        split_note = " with TEST split rows" if source_index >= 0 else ""
        self.log(
            f"Uploaded {uploaded} new row(s) to '{ws.title}'{split_note} "
            f"using {write_requests} write request(s)."
        )
        return uploaded

    def _next_append_row(self, ws):
        """Return the 1-based row number where the next append should start."""
        try:
            return len(ws.get_all_values()) + 1
        except Exception:
            return 2

    def _color_sheet_rows(self, ws, start_row_1based, end_row_1based, col_count, kind="test"):
        """Best-effort row coloring. Failures should not block data upload."""
        try:
            if start_row_1based <= 1 or end_row_1based < start_row_1based:
                return
            if kind == "test":
                bg = {"red": 1.0, "green": 0.89, "blue": 0.40}
            else:
                bg = {"red": 1.0, "green": 1.0, "blue": 1.0}
            self._spreadsheet.batch_update({
                "requests": [{
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": start_row_1based - 1,
                            "endRowIndex": end_row_1based,
                            "startColumnIndex": 0,
                            "endColumnIndex": col_count,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": bg}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                }]
            })
        except Exception as exc:
            self.log(f"Row coloring skipped for '{getattr(ws, 'title', 'sheet')}': {exc}")

    # ------------------------------------------------------------------
    # Chart creation (Sheets API v4 batchUpdate)
    # ------------------------------------------------------------------

    def _add_telemetry_charts(self):
        """
        Adds three line charts to the Telemetry sheet:
          • Temperature vs Timestamp
          • Pressure    vs Timestamp
          • Altitude    vs Timestamp
        """
        sheet_id = self._ws_tel.id
        spreadsheet_id = self._spreadsheet.id

        # Column indices (0-based) after row_id, mission_id, source, seq:
        # timestamp=4, temp=5, pressure=6, altitude=7
        chart_specs = [
            ("Temperature (°C)",   5, 0,   0),   # (title, data_col, anchor_col, anchor_row)
            ("Pressure (hPa)",     6, 7,   0),
            ("Altitude (m)",       7, 14,  0),
        ]

        requests = []
        for title, data_col, anchor_col, anchor_row in chart_specs:
            requests.append({
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": title,
                            "basicChart": {
                                "chartType": "LINE",
                                "legendPosition": "BOTTOM_LEGEND",
                                "axis": [
                                    {"position": "BOTTOM_AXIS",
                                     "title": "Timestamp"},
                                    {"position": "LEFT_AXIS",
                                     "title": title},
                                ],
                                "domains": [{
                                    "domain": {
                                        "sourceRange": {
                                            "sources": [{
                                                "sheetId": sheet_id,
                                                "startRowIndex": 0,
                                                "endRowIndex": 5000,
                                                "startColumnIndex": 4,  # timestamp column
                                                "endColumnIndex": 5,
                                            }]
                                        }
                                    }
                                }],
                                "series": [{
                                    "series": {
                                        "sourceRange": {
                                            "sources": [{
                                                "sheetId": sheet_id,
                                                "startRowIndex": 0,
                                                "endRowIndex": 5000,
                                                "startColumnIndex": data_col,
                                                "endColumnIndex": data_col + 1,
                                            }]
                                        }
                                    },
                                    "targetAxis": "LEFT_AXIS",
                                }],
                                "headerCount": 1,
                            }
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": sheet_id,
                                    "rowIndex": anchor_row + 3,
                                    "columnIndex": anchor_col + 16,
                                },
                                "widthPixels": 480,
                                "heightPixels": 300,
                            }
                        }
                    }
                }
            })

        self._spreadsheet.batch_update({"requests": requests})
        self.log("Charts added to Telemetry sheet.")


# ---------------------------------------------------------------------------
# Google Sheets configuration dialog
# ---------------------------------------------------------------------------

class SheetsConfigDialog(tk.Toplevel):
    """
    Modal dialog to enter:
      - Google Sheet URL
      - Path to service-account JSON key file

    If save_enabled is True, the caller can save these settings so the SHEETS
    button can start directly next time.
    """
    def __init__(self, parent, default_url="", default_creds="", save_enabled=True):
        super().__init__(parent)
        self.title("Google Sheets Configuration")
        self.resizable(False, False)
        self.grab_set()

        self.result = None   # will be set to (url, creds_path, save_config) on OK
        self._save_enabled = save_enabled

        pad = {"padx": 10, "pady": 6}

        ttk.Label(self, text="Google Sheet URL:").grid(
            row=0, column=0, sticky="w", **pad)
        self._url_var = tk.StringVar(value=default_url or "")
        ttk.Entry(self, textvariable=self._url_var, width=60).grid(
            row=0, column=1, **pad)

        ttk.Label(self, text="Service Account JSON:").grid(
            row=1, column=0, sticky="w", **pad)
        self._creds_var = tk.StringVar(value=default_creds or "")
        creds_frame = ttk.Frame(self)
        creds_frame.grid(row=1, column=1, sticky="w", **pad)
        ttk.Entry(creds_frame, textvariable=self._creds_var, width=50).pack(
            side="left")
        ttk.Button(creds_frame, text="Browse...",
                   command=self._browse).pack(side="left", padx=(4, 0))

        self._save_var = tk.BooleanVar(value=True)
        if save_enabled:
            ttk.Checkbutton(
                self,
                text="Remember this and start Sheets directly next time",
                variable=self._save_var,
            ).grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 4))

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Start Upload",
                   command=self._ok).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Cancel",
                   command=self.destroy).pack(side="left", padx=8)

    def _browse(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if path:
            self._creds_var.set(path)

    def _ok(self):
        url = self._url_var.get().strip()
        creds = self._creds_var.get().strip()
        if not url or not creds:
            messagebox.showerror("Error", "Both fields are required.", parent=self)
            return
        self.result = (url, creds, bool(self._save_var.get()) if self._save_enabled else False)
        self.destroy()



# ---------------------------------------------------------------------------
# Google Earth live KML bridge
# ---------------------------------------------------------------------------

class GoogleEarthLink:
    """Writes a small auto-refreshing KML feed for Google Earth Pro."""

    _REFRESH_SECONDS = 2

    def __init__(self, kml_dir, log_fn=None):
        self.log = log_fn or (lambda msg: None)
        self.base_dir = kml_dir
        os.makedirs(self.base_dir, exist_ok=True)

        self.loader_kml = os.path.join(self.base_dir, "cansat_live_loader.kml")
        self.data_kml = os.path.join(self.base_dir, "cansat_live_data.kml")
        self.path_config = os.path.join(self.base_dir, "google_earth_path.txt")
        self.google_earth_exe = self._load_saved_path() or self._find_google_earth()

        self._ground = None
        self._current = None
        self._trail = []
        self._write_loader_kml()
        self._write_data_kml()

    @staticmethod
    def _xml(text):
        text = "" if text is None else str(text)
        return (text.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace('"', "&quot;")
                    .replace("'", "&apos;"))

    @staticmethod
    def _coord(lat, lon, alt=0.0):
        return f"{float(lon):.7f},{float(lat):.7f},{float(alt or 0.0):.1f}"

    @staticmethod
    def _file_uri(path):
        return Path(path).resolve().as_uri()

    def _load_saved_path(self):
        try:
            with open(self.path_config, "r", encoding="utf-8") as f:
                path = f.read().strip()
            if path and os.path.isfile(path):
                return path
        except Exception:
            pass
        return ""

    def _find_google_earth(self):
        candidates = []

        local_appdata = os.environ.get("LOCALAPPDATA")
        program_files = os.environ.get("ProgramFiles")
        program_files_x86 = os.environ.get("ProgramFiles(x86)")

        if program_files:
            candidates.append(os.path.join(
                program_files, "Google", "Google Earth Pro", "client", "googleearth.exe"
            ))
        if program_files_x86:
            candidates.append(os.path.join(
                program_files_x86, "Google", "Google Earth Pro", "client", "googleearth.exe"
            ))
        if local_appdata:
            candidates.append(os.path.join(
                local_appdata, "Google", "Google Earth Pro", "client", "googleearth.exe"
            ))

        for exe_name in ("googleearth", "googleearthpro", "google-earth-pro"):
            found = shutil.which(exe_name)
            if found:
                candidates.append(found)

        for path in candidates:
            if path and os.path.isfile(path):
                return path
        return ""

    def set_google_earth_exe(self, path):
        if path and os.path.isfile(path):
            self.google_earth_exe = path
            try:
                with open(self.path_config, "w", encoding="utf-8") as f:
                    f.write(path)
            except Exception as exc:
                self.log(f"Google Earth path saved for this session, but not to disk: {exc}")
            self.log(f"Google Earth path set: {path}")
            return True
        return False

    def update_ground(self, lat, lon):
        self._ground = (float(lat), float(lon))
        self._write_data_kml()

    def update_position(self, lat, lon, alt=0.0, timestamp=None):
        timestamp = timestamp or datetime.now().isoformat(timespec="seconds")
        self._current = (float(lat), float(lon), float(alt or 0.0), timestamp)
        self._trail.append((float(lat), float(lon), float(alt or 0.0)))
        if len(self._trail) > 500:
            self._trail = self._trail[-500:]
        self._write_data_kml()

    def _write_loader_kml(self):
        href = self._file_uri(self.data_kml)
        content = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>CanSat Live Telemetry Loader</name>
    <NetworkLink>
      <name>CanSat live position</name>
      <open>1</open>
      <Link>
        <href>{self._xml(href)}</href>
        <refreshMode>onInterval</refreshMode>
        <refreshInterval>{self._REFRESH_SECONDS}</refreshInterval>
      </Link>
    </NetworkLink>
  </Document>
</kml>
"""
        with open(self.loader_kml, "w", encoding="utf-8") as f:
            f.write(content)

    def _write_data_kml(self):
        now = datetime.now().isoformat(timespec="seconds")
        parts = [f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>CanSat Live Telemetry</name>
    <Style id="satStyle">
      <IconStyle>
        <color>ff00ffff</color>
        <scale>1.2</scale>
        <Icon><href>http://maps.google.com/mapfiles/kml/shapes/target.png</href></Icon>
      </IconStyle>
      <LabelStyle><color>ff00ffff</color><scale>0.9</scale></LabelStyle>
    </Style>
    <Style id="groundStyle">
      <IconStyle>
        <color>ff9000ff</color>
        <scale>1.1</scale>
        <Icon><href>http://maps.google.com/mapfiles/kml/shapes/ranger_station.png</href></Icon>
      </IconStyle>
      <LabelStyle><color>ff9000ff</color><scale>0.9</scale></LabelStyle>
    </Style>
    <Style id="trailStyle">
      <LineStyle><color>ff00ffff</color><width>3</width></LineStyle>
    </Style>
    <description>Updated {self._xml(now)}</description>
"""]

        if self._ground:
            glat, glon = self._ground
            parts.append(f"""    <Placemark>
      <name>Ground Station</name>
      <styleUrl>#groundStyle</styleUrl>
      <Point><coordinates>{self._coord(glat, glon, 0)}</coordinates></Point>
    </Placemark>
""")

        if self._current:
            lat, lon, alt, timestamp = self._current
            parts.append(f"""    <Placemark id="cansat_current">
      <name>CanSat</name>
      <description>Altitude: {alt:.1f} m | Last update: {self._xml(timestamp)}</description>
      <styleUrl>#satStyle</styleUrl>
      <Point>
        <altitudeMode>absolute</altitudeMode>
        <coordinates>{self._coord(lat, lon, alt)}</coordinates>
      </Point>
    </Placemark>
    <LookAt>
      <longitude>{lon:.7f}</longitude>
      <latitude>{lat:.7f}</latitude>
      <range>900</range>
      <tilt>45</tilt>
      <heading>0</heading>
    </LookAt>
""")

        if len(self._trail) >= 2:
            coords = " ".join(self._coord(lat, lon, alt) for lat, lon, alt in self._trail)
            parts.append(f"""    <Placemark>
      <name>Flight trail</name>
      <styleUrl>#trailStyle</styleUrl>
      <LineString>
        <tessellate>1</tessellate>
        <altitudeMode>absolute</altitudeMode>
        <coordinates>{coords}</coordinates>
      </LineString>
    </Placemark>
""")

        parts.append("""  </Document>
</kml>
""")
        with open(self.data_kml, "w", encoding="utf-8") as f:
            f.write("".join(parts))

    def open(self):
        self._write_loader_kml()
        self._write_data_kml()

        if self.google_earth_exe and os.path.isfile(self.google_earth_exe):
            subprocess.Popen([self.google_earth_exe, self.loader_kml])
            return True, f"Opened Google Earth: {self.loader_kml}"

        try:
            if os.name == "nt":
                os.startfile(self.loader_kml)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.loader_kml])
            else:
                subprocess.Popen(["xdg-open", self.loader_kml])
            return True, f"Opened KML with the default application: {self.loader_kml}"
        except Exception as exc:
            return False, f"Could not open Google Earth/KML automatically: {exc}"

# ---------------------------------------------------------------------------
# Local Esri tile cache / proxy
# ---------------------------------------------------------------------------

class LocalEsriTileCache:
    """Small localhost cache for Esri World Imagery tiles.

    tkintermapview only needs a normal XYZ URL template. This helper exposes a
    local URL, saves every downloaded tile under Tile_Cache/Esri_World_Imagery,
    and can prefetch the tiles around the current map target when the map opens.
    """

    ESRI_TILE_URL = (
        "https://services.arcgisonline.com/ArcGIS/rest/services/"
        "World_Imagery/MapServer/tile/{z}/{y}/{x}"
    )

    def __init__(self, cache_dir, log_fn=None):
        self.cache_dir = os.path.join(cache_dir, "Esri_World_Imagery")
        os.makedirs(self.cache_dir, exist_ok=True)
        self.log = log_fn or (lambda msg: None)
        self._server = None
        self._thread = None
        self._lock = threading.Lock()
        self._prefetch_running = False
        self.port = None

    def start(self):
        """Start the localhost tile server if it is not already running."""
        if self._server is not None:
            return True

        cache = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_args):
                return

            def do_GET(self):
                parts = self.path.split("?", 1)[0].strip("/").split("/")
                if len(parts) != 4 or parts[0] != "esri":
                    self.send_error(404)
                    return

                try:
                    z = int(parts[1])
                    y = int(parts[2])
                    x_part = parts[3].split(".", 1)[0]
                    x = int(x_part)
                except ValueError:
                    self.send_error(400)
                    return

                data, content_type = cache.get_tile(z, x, y)
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Cache-Control", "public, max-age=31536000")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

        try:
            self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
            self.port = self._server.server_address[1]
            self._thread = threading.Thread(
                target=self._server.serve_forever,
                daemon=True,
                name="EsriTileCacheServer",
            )
            self._thread.start()
            self.log(f"Esri tile cache server started on localhost:{self.port}")
            return True
        except Exception as exc:
            self._server = None
            self.log(f"Could not start Esri tile cache server: {exc}")
            return False

    def tile_url_template(self):
        if self._server is None and not self.start():
            return self.ESRI_TILE_URL
        return f"http://127.0.0.1:{self.port}/esri/{{z}}/{{y}}/{{x}}.jpg"

    def stop(self):
        if self._server is None:
            return
        try:
            self._server.shutdown()
            self._server.server_close()
        except Exception:
            pass
        self._server = None

    def _tile_path(self, z, x, y):
        return os.path.join(self.cache_dir, str(z), str(x), f"{y}.jpg")

    def get_tile(self, z, x, y):
        path = self._tile_path(z, x, y)
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    return f.read(), "image/jpeg"
            except Exception:
                pass

        ok = self._download_tile(z, x, y)
        if ok and os.path.exists(path):
            with open(path, "rb") as f:
                return f.read(), "image/jpeg"

        return self._placeholder_tile(), "image/png"

    def _download_tile(self, z, x, y):
        path = self._tile_path(z, x, y)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        url = self.ESRI_TILE_URL.format(z=z, x=x, y=y)
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "CanSatGroundStation/1.0 (+local tile cache)",
                "Accept": "image/jpeg,image/png,*/*",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=12) as response:
                data = response.read()
            if not data:
                return False
            with self._lock:
                with open(tmp, "wb") as f:
                    f.write(data)
                os.replace(tmp, path)
            return True
        except Exception:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            return False

    def _placeholder_tile(self):
        img = Image.new("RGB", (256, 256), (8, 12, 18))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    @staticmethod
    def latlon_to_tile(lat, lon, zoom):
        lat = max(-85.05112878, min(85.05112878, float(lat)))
        lon = max(-180.0, min(180.0, float(lon)))
        lat_rad = math.radians(lat)
        n = 2 ** int(zoom)
        x = int((lon + 180.0) / 360.0 * n)
        y = int(
            (1.0 - math.log(math.tan(lat_rad) + (1.0 / math.cos(lat_rad))) / math.pi)
            / 2.0
            * n
        )
        return max(0, min(n - 1, x)), max(0, min(n - 1, y))

    def prefetch_around(self, lat, lon, zoom=17, radius=2):
        """Download a small square of tiles around a target point in background."""
        if lat is None or lon is None:
            return
        if self._prefetch_running:
            return

        self._prefetch_running = True
        thread = threading.Thread(
            target=self._prefetch_worker,
            args=(float(lat), float(lon), int(zoom), int(radius)),
            daemon=True,
            name="EsriTilePrefetch",
        )
        thread.start()

    def _prefetch_worker(self, lat, lon, zoom, radius):
        try:
            jobs = []
            # Current zoom gets a wider area. Adjacent zoom levels make zooming
            # in/out immediately after launch less likely to show blank tiles.
            for z, r in ((zoom - 1, max(1, radius - 1)), (zoom, radius), (zoom + 1, radius)):
                if z < 1 or z > 19:
                    continue
                cx, cy = self.latlon_to_tile(lat, lon, z)
                n = 2 ** z
                for dx in range(-r, r + 1):
                    for dy in range(-r, r + 1):
                        x = cx + dx
                        y = cy + dy
                        if 0 <= x < n and 0 <= y < n:
                            jobs.append((z, x, y))

            total = len(jobs)
            if total == 0:
                return
            self.log(f"Prefetching {total} Esri satellite tile(s) into cache...")
            downloaded = 0
            cached = 0
            failed = 0
            for z, x, y in jobs:
                if os.path.exists(self._tile_path(z, x, y)):
                    cached += 1
                    continue
                if self._download_tile(z, x, y):
                    downloaded += 1
                else:
                    failed += 1
            self.log(
                f"Esri tile cache ready: {cached} already cached, "
                f"{downloaded} downloaded, {failed} failed."
            )
        finally:
            self._prefetch_running = False

# ---------------------------------------------------------------------------
# Live Map Widget
# ---------------------------------------------------------------------------

class LiveMapWindow:
    """
    Standalone satellite-imagery map window for the CanSat ground station.

    Opens as an independent Toplevel so it can be moved to a second monitor.
    Closing the window just hides it; it can be re-opened from the toolbar.

    Tile source: selectable satellite/road tile servers.
    Fallback:    Pure-canvas compass-rose when tkintermapview is not installed.

    Public API
    ----------
    window.show()                        – bring window to front (create if needed)
    window.update_cansat(lat, lon, alt)  – called on every GPS packet
    window.set_ground_station(lat, lon)  – set / move the GND marker
    """

    _BG        = "#020208"
    _CYAN      = "#00FFE5"
    _MAGENTA   = "#FF0090"
    _DIM       = "#004D47"
    _TRAIL_MAX = 200

    # Satellite tile servers. Esri is the default because it is actual imagery
    # and does not require a Google Maps API key for this Tkinter widget.
    _TILE_SERVERS = {
        "Esri Satellite": (
            "https://services.arcgisonline.com/ArcGIS/rest/services/"
            "World_Imagery/MapServer/tile/{z}/{y}/{x}",
            19,
        ),
        "Google Satellite": (
            "https://mt0.google.com/vt/lyrs=s&hl=en&x={x}&y={y}&z={z}&s=Ga",
            22,
        ),
        "Google Hybrid": (
            "https://mt0.google.com/vt/lyrs=y&hl=en&x={x}&y={y}&z={z}&s=Ga",
            22,
        ),
        "OpenStreetMap": (
            "https://a.tile.openstreetmap.org/{z}/{x}/{y}.png",
            19,
        ),
    }

    def __init__(self, root, app):
        self._root = root
        self.app   = app

        self._cansat_lat    = None
        self._cansat_lon    = None
        self._gs_lat        = None
        self._gs_lon        = None
        self._trail         = []
        self._trail_line    = None
        self._cansat_marker = None
        self._gs_marker     = None
        self._bearing_line  = None

        self._win = None   # tk.Toplevel, created lazily
        self._tile_cache = getattr(app, "_tile_cache", None)
        self._last_prefetch_target = None

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def show(self):
        """Create the window if it does not exist, then bring it to front."""
        self._tile_cache = getattr(self.app, "_tile_cache", self._tile_cache)
        if self._win is None or not self._win.winfo_exists():
            self._create_window()
        else:
            self._win.deiconify()
            self._win.lift()
            self._maybe_prefetch_esri_tiles()

    def update_cansat(self, lat: float, lon: float, alt: float = 0.0):
        self._cansat_lat = lat
        self._cansat_lon = lon
        self._trail.append((lat, lon))
        if len(self._trail) > self._TRAIL_MAX:
            self._trail.pop(0)

        if self._win is None or not self._win.winfo_exists():
            return   # window not open – buffer data silently

        self._refresh_info()
        if _HAS_MAPVIEW:
            self._refresh_map()
        else:
            self._redraw_fallback()

    def set_ground_station(self, lat: float, lon: float):
        self._gs_lat = lat
        self._gs_lon = lon

        if self._win is None or not self._win.winfo_exists():
            return

        self._gs_ll_var.set(f"{lat:.5f}, {lon:.5f}")
        self._gs_lat_entry.delete(0, "end")
        self._gs_lat_entry.insert(0, str(lat))
        self._gs_lon_entry.delete(0, "end")
        self._gs_lon_entry.insert(0, str(lon))

        ge = getattr(self.app, "_google_earth", None)
        if ge is not None:
            ge.update_ground(lat, lon)

        if _HAS_MAPVIEW:
            self._place_gs_marker()
            if self._cansat_lat is None:
                self._map.set_position(lat, lon)
            self._maybe_prefetch_esri_tiles()
        else:
            self._redraw_fallback()

    # ------------------------------------------------------------------
    # Window creation
    # ------------------------------------------------------------------

    def _create_window(self):
        self._win = tk.Toplevel(self._root)
        self._win.title("CanSat — Live Satellite Map")
        self._win.geometry("960x720")
        self._win.configure(bg=self._BG)
        self._win.protocol("WM_DELETE_WINDOW", self._win.withdraw)

        # StringVars must be re-created when the window is created
        self._dist_var    = tk.StringVar(value="—")
        self._bearing_var = tk.StringVar(value="—")
        self._csat_ll_var = tk.StringVar(value="—")
        self._gs_ll_var   = tk.StringVar(value="set below")
        self._tile_var    = tk.StringVar(value="Esri Satellite")
        self._tile_cache_var = tk.StringVar(value="cache: ready")

        self._build_ui()

        # Restore existing state
        if self._gs_lat is not None:
            self.set_ground_station(self._gs_lat, self._gs_lon)
        if self._cansat_lat is not None:
            self.update_cansat(self._cansat_lat, self._cansat_lon)
        self._maybe_prefetch_esri_tiles()

    def _build_ui(self):
        win = self._win

        # ── title bar ─────────────────────────────────────────────────
        title_bar = tk.Frame(win, bg="#05080F", pady=6)
        title_bar.pack(fill="x")
        tk.Label(title_bar, text="🛰  CANSAT  LIVE SATELLITE MAP",
                 bg="#05080F", fg=self._CYAN,
                 font=("Consolas", 13, "bold")).pack(side="left", padx=14)

        # ── info strip ────────────────────────────────────────────────
        info = tk.Frame(win, bg="#0A0E14", pady=5)
        info.pack(fill="x")

        def _lbl(text, col):
            tk.Label(info, text=text, bg="#0A0E14", fg=self._DIM,
                     font=("Consolas", 9)).grid(
                         row=0, column=col, sticky="w", padx=(10, 2))

        def _val(var, col, color=None):
            tk.Label(info, textvariable=var, bg="#0A0E14",
                     fg=color or self._CYAN,
                     font=("Consolas", 10, "bold"), width=20,
                     anchor="w").grid(row=0, column=col, sticky="w", padx=2)

        _lbl("DISTANCE",  0); _val(self._dist_var,    1)
        _lbl("BEARING",   2); _val(self._bearing_var, 3)
        _lbl("CANSAT",    4); _val(self._csat_ll_var, 5)
        _lbl("GND STN",   6); _val(self._gs_ll_var,   7, color=self._MAGENTA)

        # ── ground-station coordinate entry ───────────────────────────
        gs_bar = tk.Frame(win, bg="#0A0E14", pady=4)
        gs_bar.pack(fill="x")

        tk.Label(gs_bar, text="GND LAT", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).pack(side="left", padx=(12, 2))
        self._gs_lat_entry = tk.Entry(gs_bar, width=13,
                                      font=("Consolas", 9),
                                      bg="#11151C", fg=self._MAGENTA,
                                      insertbackground=self._MAGENTA,
                                      relief="flat")
        self._gs_lat_entry.pack(side="left", padx=(0, 8))

        tk.Label(gs_bar, text="GND LON", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).pack(side="left", padx=(0, 2))
        self._gs_lon_entry = tk.Entry(gs_bar, width=13,
                                      font=("Consolas", 9),
                                      bg="#11151C", fg=self._MAGENTA,
                                      insertbackground=self._MAGENTA,
                                      relief="flat")
        self._gs_lon_entry.pack(side="left", padx=(0, 8))

        tk.Button(gs_bar, text="SET GND STATION",
                  font=("Consolas", 9, "bold"),
                  bg=self._MAGENTA, fg="#000",
                  activebackground="#FF33AA", relief="flat", padx=10,
                  command=self._on_set_gs).pack(side="left", padx=4)

        tk.Label(gs_bar, text="MAP", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).pack(side="left", padx=(18, 2))
        self._tile_combo = ttk.Combobox(
            gs_bar,
            textvariable=self._tile_var,
            values=list(self._TILE_SERVERS.keys()),
            width=18,
            state="readonly",
        )
        self._tile_combo.pack(side="left", padx=(0, 8))
        self._tile_combo.bind("<<ComboboxSelected>>", self._on_tile_change)

        tk.Button(gs_bar, text="PREFETCH ESRI",
                  font=("Consolas", 8, "bold"),
                  bg="#264653", fg="white",
                  activebackground="#2A9D8F", relief="flat", padx=8,
                  command=self._prefetch_esri_tiles).pack(side="left", padx=4)

        tk.Label(gs_bar, textvariable=self._tile_cache_var,
                 bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8), width=14, anchor="w").pack(side="left", padx=(0, 6))

        tk.Button(gs_bar, text="OPEN GOOGLE MAPS SAT",
                  font=("Consolas", 8, "bold"),
                  bg="#2D6A4F", fg="white",
                  activebackground="#40916C", relief="flat", padx=8,
                  command=self._open_google_maps_satellite).pack(side="left", padx=4)

        tk.Button(gs_bar, text="OPEN GOOGLE EARTH",
                  font=("Consolas", 8, "bold"),
                  bg="#3A0CA3", fg="white",
                  activebackground="#5F37D6", relief="flat", padx=8,
                  command=self._open_google_earth).pack(side="left", padx=4)

        tk.Button(gs_bar, text="GE PATH...",
                  font=("Consolas", 8, "bold"),
                  bg="#222831", fg="white",
                  activebackground="#393E46", relief="flat", padx=8,
                  command=self._choose_google_earth_path).pack(side="left", padx=4)

        # ── map / fallback canvas ──────────────────────────────────────
        map_frame = tk.Frame(win, bg=self._BG)
        map_frame.pack(fill="both", expand=True)

        if _HAS_MAPVIEW:
            self._map = tkintermapview.TkinterMapView(
                map_frame, corner_radius=0)
            self._map.pack(fill="both", expand=True)
            self._apply_tile_server()
            self._map.set_zoom(17)
        else:
            self._map = None
            self._canvas = tk.Canvas(map_frame, bg=self._BG,
                                     highlightthickness=0)
            self._canvas.pack(fill="both", expand=True)
            self._canvas.bind("<Configure>",
                              lambda e: self._redraw_fallback())
            self._root.after(120, self._redraw_fallback)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _on_set_gs(self):
        try:
            lat = float(self._gs_lat_entry.get())
            lon = float(self._gs_lon_entry.get())
        except ValueError:
            return
        if hasattr(self.app, "set_ground_station_location"):
            self.app.set_ground_station_location(lat, lon, source="manual entry")
        else:
            self.set_ground_station(lat, lon)

    def _on_tile_change(self, _event=None):
        self._apply_tile_server()
        if self._cansat_lat is not None:
            self._refresh_map()
        elif self._gs_lat is not None and _HAS_MAPVIEW:
            self._map.set_position(self._gs_lat, self._gs_lon)

    def _apply_tile_server(self):
        if not _HAS_MAPVIEW or self._map is None:
            return
        name = self._tile_var.get() or "Esri Satellite"
        tile_url, max_zoom = self._TILE_SERVERS.get(
            name, self._TILE_SERVERS["Esri Satellite"]
        )
        if name == "Esri Satellite" and self._tile_cache is not None:
            tile_url = self._tile_cache.tile_url_template()
            if hasattr(self, "_tile_cache_var"):
                self._tile_cache_var.set("cache: localhost")
        elif hasattr(self, "_tile_cache_var"):
            self._tile_cache_var.set("cache: off")
        self._map.set_tile_server(tile_url, max_zoom=max_zoom)

    def _target_position(self):
        if self._cansat_lat is not None and self._cansat_lon is not None:
            return self._cansat_lat, self._cansat_lon
        if self._gs_lat is not None and self._gs_lon is not None:
            return self._gs_lat, self._gs_lon
        return None

    def _prefetch_esri_tiles(self):
        if self._tile_cache is None:
            return
        pos = self._target_position()
        if not pos:
            return
        lat, lon = pos
        zoom = 17
        try:
            if _HAS_MAPVIEW and self._map is not None:
                zoom = int(self._map.zoom)
        except Exception:
            pass
        self._tile_cache_var.set("cache: warming")
        self._tile_cache.prefetch_around(lat, lon, zoom=zoom, radius=2)
        if self._win is not None and self._win.winfo_exists():
            self._win.after(3000, lambda: self._tile_cache_var.set("cache: localhost"))
        self._last_prefetch_target = (round(lat, 5), round(lon, 5), zoom)

    def _maybe_prefetch_esri_tiles(self):
        if self._tile_var.get() != "Esri Satellite" or self._tile_cache is None:
            return
        pos = self._target_position()
        if not pos:
            return
        lat, lon = pos
        zoom = 17
        try:
            if _HAS_MAPVIEW and self._map is not None:
                zoom = int(self._map.zoom)
        except Exception:
            pass
        target = (round(lat, 5), round(lon, 5), zoom)
        if target != self._last_prefetch_target:
            self._tile_cache_var.set("cache: warming")
            self._tile_cache.prefetch_around(lat, lon, zoom=zoom, radius=2)
            if self._win is not None and self._win.winfo_exists():
                self._win.after(3000, lambda: self._tile_cache_var.set("cache: localhost"))
            self._last_prefetch_target = target

    def _open_google_maps_satellite(self):
        pos = self._target_position()
        if not pos:
            messagebox.showinfo(
                "No coordinates yet",
                "Set the ground station or wait for a GPS fix first.",
                parent=self._win,
            )
            return
        lat, lon = pos
        url = f"https://www.google.com/maps/@{lat:.7f},{lon:.7f},18z/data=!3m1!1e3"
        webbrowser.open(url)

    def _open_google_earth(self):
        ge = getattr(self.app, "_google_earth", None)
        if ge is None:
            messagebox.showerror(
                "Google Earth",
                "Google Earth exporter is not initialized.",
                parent=self._win,
            )
            return
        if self._gs_lat is not None:
            ge.update_ground(self._gs_lat, self._gs_lon)
        if self._cansat_lat is not None:
            ge.update_position(self._cansat_lat, self._cansat_lon, 0.0)
        ok, msg = ge.open()
        self.app.log(msg)
        if not ok:
            messagebox.showerror("Google Earth", msg, parent=self._win)

    def _choose_google_earth_path(self):
        from tkinter import filedialog
        ge = getattr(self.app, "_google_earth", None)
        if ge is None:
            return
        path = filedialog.askopenfilename(
            parent=self._win,
            title="Select Google Earth Pro executable",
            filetypes=[("Google Earth executable", "googleearth.exe"),
                       ("Executable files", "*.exe"),
                       ("All files", "*.*")],
        )
        if not path:
            return
        if ge.set_google_earth_exe(path):
            messagebox.showinfo("Google Earth", "Google Earth path saved.", parent=self._win)
        else:
            messagebox.showerror("Google Earth", "That file does not exist.", parent=self._win)

    def _haversine(self, lat1, lon1, lat2, lon2):
        """Return (distance_m, bearing_deg)."""
        R = 6_371_000.0
        phi1, phi2 = math.radians(lat1), math.radians(lat2)
        dphi = math.radians(lat2 - lat1)
        dlam = math.radians(lon2 - lon1)
        a = (math.sin(dphi / 2) ** 2
             + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2)
        dist = 2 * R * math.asin(math.sqrt(a))
        y = math.sin(dlam) * math.cos(phi2)
        x = (math.cos(phi1) * math.sin(phi2)
             - math.sin(phi1) * math.cos(phi2) * math.cos(dlam))
        bearing = (math.degrees(math.atan2(y, x)) + 360) % 360
        return dist, bearing

    def _refresh_info(self):
        if self._cansat_lat is None:
            return
        self._csat_ll_var.set(
            f"{self._cansat_lat:.5f}, {self._cansat_lon:.5f}")

        if self._gs_lat is not None:
            dist, brg = self._haversine(
                self._gs_lat, self._gs_lon,
                self._cansat_lat, self._cansat_lon)
            self._dist_var.set(
                f"{dist / 1000:.2f} km" if dist >= 1000 else f"{dist:.1f} m")
            dirs = ["N", "NE", "E", "SE", "S", "SW", "W", "NW", "N"]
            self._bearing_var.set(
                f"{brg:.1f}°  {dirs[round(brg / 45) % 8]}")

    # ── tkintermapview helpers ─────────────────────────────────────────

    def _refresh_map(self):
        if not _HAS_MAPVIEW or self._cansat_lat is None:
            return
        if self._cansat_marker:
            self._cansat_marker.delete()
        if self._bearing_line:
            self._bearing_line.delete()
        if self._trail_line:
            self._trail_line.delete()

        if len(self._trail) >= 2:
            self._trail_line = self._map.set_path(
                self._trail, color=self._CYAN, width=2)

        self._cansat_marker = self._map.set_marker(
            self._cansat_lat, self._cansat_lon,
            text="CanSat",
            marker_color_circle=self._CYAN,
            marker_color_outside=self._CYAN,
            text_color=self._CYAN,
        )

        if self._gs_lat is not None:
            self._place_gs_marker()
            self._bearing_line = self._map.set_path(
                [(self._gs_lat, self._gs_lon),
                 (self._cansat_lat, self._cansat_lon)],
                color=self._MAGENTA, width=2)

        self._map.set_position(self._cansat_lat, self._cansat_lon)
        self._maybe_prefetch_esri_tiles()

    def _place_gs_marker(self):
        if not _HAS_MAPVIEW or self._gs_lat is None:
            return
        if self._gs_marker:
            self._gs_marker.delete()
        self._gs_marker = self._map.set_marker(
            self._gs_lat, self._gs_lon,
            text="GND",
            marker_color_circle=self._MAGENTA,
            marker_color_outside=self._MAGENTA,
            text_color=self._MAGENTA,
        )

    # ── Fallback canvas compass-rose ──────────────────────────────────

    def _redraw_fallback(self):
        c = self._canvas
        c.delete("all")
        w = c.winfo_width()  or 600
        h = c.winfo_height() or 560
        cx, cy = w / 2, h / 2
        r = min(cx, cy) - 40

        c.create_oval(cx - r, cy - r, cx + r, cy + r,
                      outline=self._DIM, width=2)

        for deg, lbl in [(0, "N"), (90, "E"), (180, "S"), (270, "W")]:
            rad = math.radians(deg - 90)
            x1 = cx + (r - 14) * math.cos(rad)
            y1 = cy + (r - 14) * math.sin(rad)
            x2 = cx + r * math.cos(rad)
            y2 = cy + r * math.sin(rad)
            c.create_line(x1, y1, x2, y2, fill=self._DIM, width=2)
            xl = cx + (r + 16) * math.cos(rad)
            yl = cy + (r + 16) * math.sin(rad)
            c.create_text(xl, yl, text=lbl, fill=self._CYAN,
                          font=("Consolas", 11, "bold"))

        # Ground-station dot at centre
        c.create_oval(cx - 7, cy - 7, cx + 7, cy + 7,
                      fill=self._MAGENTA, outline="")
        c.create_text(cx + 16, cy, text="GND",
                      fill=self._MAGENTA,
                      font=("Consolas", 9, "bold"), anchor="w")

        if self._cansat_lat is None or self._gs_lat is None:
            c.create_text(cx, cy + r + 30,
                          text="Waiting for GPS fix…",
                          fill=self._DIM, font=("Consolas", 10))
            return

        dist, bearing = self._haversine(
            self._gs_lat, self._gs_lon,
            self._cansat_lat, self._cansat_lon)

        scale = min(dist / 2000.0, 1.0)
        rad = math.radians(bearing - 90)
        sat_x = cx + scale * (r - 22) * math.cos(rad)
        sat_y = cy + scale * (r - 22) * math.sin(rad)

        # Trail dots
        if len(self._trail) >= 2:
            trail_pts = self._trail[-60:]
            for i, (tlat, tlon) in enumerate(trail_pts):
                td, tb = self._haversine(
                    self._gs_lat, self._gs_lon, tlat, tlon)
                ts = min(td / 2000.0, 1.0)
                tr = math.radians(tb - 90)
                tx = cx + ts * (r - 22) * math.cos(tr)
                ty = cy + ts * (r - 22) * math.sin(tr)
                fade = int(60 + 180 * i / len(trail_pts))
                color = f"#00{fade:02x}{int(fade*0.88):02x}"
                c.create_oval(tx - 2, ty - 2, tx + 2, ty + 2,
                              fill=color, outline="")

        # Bearing dashed line + arrowhead
        c.create_line(cx, cy, sat_x, sat_y,
                      fill=self._MAGENTA, width=1, dash=(5, 3))
        angle = math.atan2(sat_y - cy, sat_x - cx)
        for sign in (+1, -1):
            ax_ = sat_x - 10 * math.cos(angle - sign * math.pi / 6)
            ay_ = sat_y - 10 * math.sin(angle - sign * math.pi / 6)
            c.create_line(sat_x, sat_y, ax_, ay_,
                          fill=self._MAGENTA, width=2)

        # CanSat dot
        c.create_oval(sat_x - 8, sat_y - 8, sat_x + 8, sat_y + 8,
                      fill=self._CYAN, outline="white", width=1)
        c.create_text(sat_x, sat_y - 18, text="SAT",
                      fill=self._CYAN, font=("Consolas", 9, "bold"))


class GroundStationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CanSat Ground Station")
        self.root.geometry("1460x900")
        self.root.configure(bg="#0F1115")

        base_dir = os.path.dirname(os.path.abspath(__file__))

        # All generated files now live under one parent folder.
        # This keeps CSVs, raw logs, preview images, and Google Earth KML files
        # separated instead of mixing everything in one logs directory.
        self.output_root = os.path.join(base_dir, "CanSat_GroundStation_Data")
        self.csv_dir = os.path.join(self.output_root, "CSVs")
        self.kml_dir = os.path.join(self.output_root, "KMLs")
        self.preview_dir = os.path.join(self.output_root, "Previews")
        self.raw_log_dir = os.path.join(self.output_root, "Raw_Logs")
        self.tile_cache_dir = os.path.join(self.output_root, "Tile_Cache")
        self.secrets_dir = os.path.join(self.output_root, "Secrets")
        self.sheets_config_path = os.path.join(self.secrets_dir, "sheets_config.json")

        for folder in (
            self.output_root,
            self.csv_dir,
            self.kml_dir,
            self.preview_dir,
            self.raw_log_dir,
            self.tile_cache_dir,
            self.secrets_dir,
        ):
            os.makedirs(folder, exist_ok=True)

        # Backwards-compatible name used by older helper code / comments.
        self.log_dir = self.output_root

        self.telemetry_csv = os.path.join(self.csv_dir, "telemetry.csv")
        self.gps_csv = os.path.join(self.csv_dir, "gps.csv")
        self.telemetry_test_csv = os.path.join(self.csv_dir, "telemetry_TEST.csv")
        self.gps_test_csv = os.path.join(self.csv_dir, "gps_TEST.csv")
        self.raw_log = os.path.join(self.raw_log_dir, "raw_packets.log")
        self._init_mission_state()

        self.serial_port = None
        self.serial_thread = None
        self.running = False
        self.rx_queue = queue.Queue()
        self.sat_started = False

        self.packet_count = 0
        self.gps_packet_count = 0
        self.last_packet_time = None
        self.current_preview_photo = None
        self.image_buffers = {}

        self.roll = 0.0
        self.pitch = 0.0
        self.vx = 0.0
        self.vy = 0.0

        self.last_seq = None
        self.telemetry_history = []

        self._sheets_uploader = None   # GoogleSheetsUploader, set when user enables it
        self._live_map        = None   # LiveMapWindow, built in _build_ui
        self._google_earth    = None   # GoogleEarthLink, writes live KML files
        self._tile_cache      = None   # LocalEsriTileCache, serves Esri imagery from disk
        # Default ground-station position (editable in the map widget)
        self._gs_lat_default  = 59.4370   # Tallinn, adjust as needed
        self._gs_lon_default  = 24.7536

        self._test_mode_running = False
        self._test_seq = 0
        self._test_start_time = None
        self._test_timer_id = None
        self._test_peak_alt = 0.0
        self._test_preview_last = -999

        self._build_styles()
        self._build_ui()
        self.log(f"Output folder: {self.output_root}")
        self.log(f"CSV folder: {self.csv_dir}")
        self.log(f"Real CSVs: telemetry.csv, gps.csv")
        self.log(f"Test CSVs: telemetry_TEST.csv, gps_TEST.csv")
        self.log(f"KML folder: {self.kml_dir}")
        self.log(f"Tile cache folder: {self.tile_cache_dir}")
        self.log(f"Sheets config: {self.sheets_config_path}")
        self.log(f"Mission ID: {self.mission_id}")
        self._tile_cache = LocalEsriTileCache(self.tile_cache_dir, log_fn=self.log)
        self._tile_cache.start()
        self._google_earth = GoogleEarthLink(self.kml_dir, log_fn=self.log)
        self._google_earth.update_ground(self._gs_lat_default, self._gs_lon_default)
        self.refresh_ports()
        self.root.after(100, self.process_queue)
        self.root.after(500, self.update_link_status)

        self.viz_window = TelemetryVizWindow(self)

    def _build_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background="#0F1115")
        style.configure("TLabelframe", background="#151922", foreground="#E6EAF2")
        style.configure("TLabelframe.Label", background="#151922", foreground="#E6EAF2",
                        font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", background="#0F1115", foreground="#E6EAF2",
                        font=("Segoe UI", 10))
        style.configure("Header.TLabel", background="#0F1115", foreground="#7FDBFF",
                        font=("Segoe UI", 11, "bold"))
        style.configure("Value.TLabel", background="#151922", foreground="#FFFFFF",
                        font=("Consolas", 12, "bold"))
        style.configure("TButton", font=("Segoe UI", 10))

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="COM Port:", style="Header.TLabel").pack(side="left")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(top, textvariable=self.port_var, width=18, state="readonly")
        self.port_combo.pack(side="left", padx=6)

        ttk.Button(top, text="Refresh", command=self.refresh_ports).pack(side="left", padx=4)

        ttk.Label(top, text="Baud:", style="Header.TLabel").pack(side="left", padx=(14, 0))
        self.baud_var = tk.StringVar(value="57600")
        ttk.Entry(top, textvariable=self.baud_var, width=10).pack(side="left", padx=6)

        self.connect_btn = ttk.Button(top, text="Connect", command=self.toggle_connection)
        self.connect_btn.pack(side="left", padx=10)

        self.start_btn = tk.Button(
            top,
            text="▶  START SAT",
            bg="#1B5E20",
            fg="white",
            activebackground="#2E7D32",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self.start_satellite,
        )
        self.start_btn.pack(side="left", padx=6)

        self.stop_btn = tk.Button(
            top,
            text="■  STOP SAT",
            bg="#8B5E00",
            fg="white",
            activebackground="#A66F00",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self.stop_satellite,
        )
        self.stop_btn.pack(side="left", padx=6)

        self._sheets_btn = tk.Button(
            top,
            text="☁  SHEETS",
            bg="#1A237E",
            fg="white",
            activebackground="#283593",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self.toggle_sheets_upload,
        )
        self._sheets_btn.pack(side="left", padx=6)

        tk.Button(
            top,
            text="🗺  MAP",
            bg="#1B4332",
            fg="white",
            activebackground="#2D6A4F",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self._open_map_window,
        ).pack(side="left", padx=6)

        tk.Button(
            top,
            text="🌍  EARTH",
            bg="#3A0CA3",
            fg="white",
            activebackground="#5F37D6",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self._open_google_earth,
        ).pack(side="left", padx=6)

        self._test_btn = tk.Button(
            top,
            text="TEST SIM",
            bg="#6A1B9A",
            fg="white",
            activebackground="#8E24AA",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            command=self.toggle_test_mode,
        )
        self._test_btn.pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Disconnected")
        self.packet_var = tk.StringVar(value="Packets: 0")
        self.last_rx_var = tk.StringVar(value="Last RX: -")
        self.link_var = tk.StringVar(value="Link: idle")
        self.sat_status_var = tk.StringVar(value="SAT: —")

        ttk.Label(top, textvariable=self.status_var, style="Header.TLabel").pack(side="left", padx=(20, 10))
        ttk.Label(top, textvariable=self.packet_var).pack(side="left", padx=10)
        ttk.Label(top, textvariable=self.last_rx_var).pack(side="left", padx=10)
        ttk.Label(top, textvariable=self.link_var, style="Header.TLabel").pack(side="left", padx=10)
        ttk.Label(top, textvariable=self.sat_status_var, style="Header.TLabel").pack(side="left", padx=10)

        main = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(main)
        right.pack(side="right", fill="y", padx=(10, 0))

        cards = ttk.Frame(left)
        cards.pack(fill="x")

        self.card_vars = {}
        card_items = [
            ("seq", "SEQ"),
            ("temp", "TEMP °C"),
            ("pressure", "PRESS hPa"),
            ("alt", "ALT m"),
            ("peak_alt", "PEAK ALT"),
            ("deploy", "DEPLOY"),
            ("ax", "ACC X"),
            ("ay", "ACC Y"),
            ("az", "ACC Z"),
            ("gx", "GYRO X"),
            ("gy", "GYRO Y"),
            ("gz", "GYRO Z"),
            ("gps_lat", "GPS LAT"),
            ("gps_lon", "GPS LON"),
            ("gps_alt", "GNSS ALT"),
            ("gps_sats", "GPS SATS"),
        ]

        for i, (key, label) in enumerate(card_items):
            frame = ttk.LabelFrame(cards, text=label, padding=10)
            frame.grid(row=i // 8, column=i % 8, padx=5, pady=5, sticky="nsew")
            var = tk.StringVar(value="-")
            ttk.Label(frame, textvariable=var, style="Value.TLabel", width=12).pack()
            self.card_vars[key] = var

        for i in range(8):
            cards.columnconfigure(i, weight=1)

        ts_frame = ttk.LabelFrame(left, text="Timestamp", padding=10)
        ts_frame.pack(fill="x", pady=(8, 8))
        self.timestamp_var = tk.StringVar(value="-")
        ttk.Label(ts_frame, textvariable=self.timestamp_var, style="Value.TLabel").pack(anchor="w")

        lower = ttk.Frame(left)
        lower.pack(fill="both", expand=True)

        att_frame = ttk.LabelFrame(lower, text="Attitude Indicator", padding=10)
        att_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.attitude = AttitudeIndicator(att_frame, self, width=420, height=420)
        self.attitude.pack(expand=True, fill="both")
        self.attitude.draw_indicator(0.0, 0.0)

        preview_frame = ttk.LabelFrame(lower, text="Latest Preview", padding=10)
        preview_frame.pack(side="right", fill="both", expand=True)

        self.preview_label = ttk.Label(preview_frame, text="No preview yet",
                                       anchor="center")
        self.preview_label.pack(fill="both", expand=True)

        # Live map window is opened via the toolbar button; just initialise it
        self._live_map = LiveMapWindow(self.root, self)
        self._live_map.set_ground_station(self._gs_lat_default,
                                          self._gs_lon_default)

        console_frame = ttk.LabelFrame(right, text="Console", padding=10)
        console_frame.pack(fill="both", expand=True)

        self.console = tk.Text(
            console_frame,
            width=42,
            wrap="word",
            bg="#11151C",
            fg="#E6EAF2",
            insertbackground="white",
            relief="flat",
            font=("Consolas", 10)
        )
        self.console.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(console_frame, orient="vertical", command=self.console.yview)
        yscroll.pack(side="right", fill="y")
        self.console.configure(yscrollcommand=yscroll.set)

    def log(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        self.console.insert("end", f"[{ts}] {text}\n")
        self.console.see("end")

    def refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo["values"] = ports
        if ports and not self.port_var.get():
            self.port_var.set(ports[0])

    def toggle_connection(self):
        if self.running:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        port = self.port_var.get().strip()
        if not port:
            messagebox.showerror("Error", "Select a COM port.")
            return

        try:
            baud = int(self.baud_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Invalid baud rate.")
            return

        try:
            self.serial_port = serial.Serial(port, baud, timeout=1)
        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
            return

        self.running = True
        self.serial_thread = threading.Thread(target=self.serial_reader, daemon=True)
        self.serial_thread.start()

        self.status_var.set(f"Connected: {port} @ {baud}")
        self.connect_btn.config(text="Disconnect")
        self.link_var.set("Link: active")
        self.log(f"Connected to {port} @ {baud}")

    def disconnect(self):
        self.running = False
        try:
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
        except Exception:
            pass

        self.status_var.set("Disconnected")
        self.connect_btn.config(text="Connect")
        self.link_var.set("Link: idle")
        self.log("Disconnected")

    def start_satellite(self):
        if not self.running or not self.serial_port or not self.serial_port.is_open:
            messagebox.showerror("Error", "Not connected to radio.")
            return
        try:
            for _ in range(3):
                self.serial_port.write(b"CMD,START\n")
            self.log("CMD,START sent — waiting for satellite confirmation")
            self.sat_status_var.set("SAT: STARTING…")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def stop_satellite(self):
        if not self.running or not self.serial_port or not self.serial_port.is_open:
            messagebox.showerror("Error", "Not connected to radio.")
            return

        ok = messagebox.askyesno(
            "Confirm STOP SAT",
            "Send CMD,STOP to the satellite?\n\nThis should stop logging/saving remotely."
        )
        if not ok:
            return

        try:
            for _ in range(3):
                self.serial_port.write(b"CMD,STOP\n")
            self.log("CMD,STOP sent")
            self.sat_status_var.set("SAT: STOPPING…")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def handle_status(self, line):
        parts = line.split(",", 1)
        if len(parts) < 2:
            return
        status = parts[1].strip()

        if status == "READY":
            self.sat_status_var.set("SAT: STANDBY")
            self.start_btn.config(state="normal", bg="#1B5E20", fg="white")
            self.log("Satellite is standing by — press START SAT to begin")

        elif status == "RUNNING":
            self.sat_status_var.set("SAT: RUNNING ●")
            self.sat_started = True
            self.start_btn.config(state="disabled", bg="#424242", fg="#888888")
            self.log("Satellite started — recording and transmitting")

        elif status in ("STOPPED", "IDLE"):
            self.sat_status_var.set("SAT: STOPPED")
            self.start_btn.config(state="normal", bg="#1B5E20", fg="white")
            self.log("Satellite stopped logging/transmitting")

        else:
            self.log(f"SAT status: {status}")

    def serial_reader(self):
        while self.running:
            try:
                line = self.serial_port.readline().decode(errors="ignore").strip()
                if line:
                    self.rx_queue.put(line)
            except Exception as e:
                self.rx_queue.put(("ERROR", str(e)))
                break

    def process_queue(self):
        try:
            while True:
                item = self.rx_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "ERROR":
                    self.log(f"Serial error: {item[1]}")
                    self.disconnect()
                    break
                self.handle_line(item)
        except queue.Empty:
            pass

        self.root.after(100, self.process_queue)

    def update_link_status(self):
        if self.last_packet_time is None:
            self.link_var.set("Link: idle")
        else:
            age = (datetime.now() - self.last_packet_time).total_seconds()
            if age < 2:
                self.link_var.set("Link: good")
            elif age < 5:
                self.link_var.set("Link: stale")
            else:
                self.link_var.set("Link: lost")
        self.root.after(500, self.update_link_status)

    def toggle_test_mode(self):
        """Start or stop the built-in simulator for UI and logging tests."""
        if self._test_mode_running:
            self.stop_test_mode()
        else:
            self.start_test_mode()

    def start_test_mode(self):
        """Feed fake telemetry and GPS packets through the normal handlers."""
        self._test_mode_running = True
        self._test_seq = 0
        self._test_peak_alt = 0.0
        self._test_preview_last = -999
        self._test_start_time = datetime.now()
        self.last_seq = None
        self.sat_started = True
        self.sat_status_var.set("SAT: TEST MODE")
        if hasattr(self, "_test_btn"):
            self._test_btn.config(text="STOP TEST", bg="#B71C1C", activebackground="#D32F2F")
        self.log("Test simulator started. Fake telemetry/GPS will be written to CSVs and displayed like real packets.")
        self._run_test_tick()

    def stop_test_mode(self):
        """Stop the built-in simulator."""
        self._test_mode_running = False
        if self._test_timer_id is not None:
            try:
                self.root.after_cancel(self._test_timer_id)
            except Exception:
                pass
            self._test_timer_id = None
        if hasattr(self, "_test_btn"):
            self._test_btn.config(text="TEST SIM", bg="#6A1B9A", activebackground="#8E24AA")
        self.sat_status_var.set("SAT: TEST STOPPED")
        self.log("Test simulator stopped.")

    def _run_test_tick(self):
        if not self._test_mode_running:
            return

        now = datetime.now()
        if self._test_start_time is None:
            self._test_start_time = now
        t = (now - self._test_start_time).total_seconds()
        self._test_seq += 1
        seq = self._test_seq
        timestamp = now.isoformat(timespec="seconds")

        # Smooth looped flight profile. Altitude climbs and descends, while
        # temperature and pressure follow altitude. This keeps charts moving.
        phase = (2.0 * math.pi * (t % 90.0)) / 90.0
        alt = 85.0 + 360.0 * (0.5 - 0.5 * math.cos(phase)) + 8.0 * math.sin(t * 0.7)
        alt = max(0.0, alt)
        self._test_peak_alt = max(self._test_peak_alt, alt)
        temp = 20.0 - 0.0065 * alt + 1.1 * math.sin(t / 8.0)
        pressure = 1013.25 * max(0.01, (1.0 - alt / 44330.0)) ** 5.255

        # Fake IMU values calculated from target roll/pitch so both the 2-D
        # attitude indicator and the 3-D orientation tab move continuously.
        target_roll = 24.0 * math.sin(t / 4.5)
        target_pitch = 14.0 * math.sin(t / 6.5)
        rr = math.radians(target_roll)
        pp = math.radians(target_pitch)
        ax = -math.sin(pp)
        ay = math.sin(rr) * math.cos(pp)
        az = math.cos(rr) * math.cos(pp)
        gx = 35.0 * math.cos(t / 4.5)
        gy = 22.0 * math.cos(t / 6.5)
        gz = 8.0 * math.sin(t / 5.0)

        descending = math.sin(phase) < -0.15
        deployed = "1" if descending and alt < 230.0 else "0"
        deploy_reason = "TEST_DESCENT" if deployed == "1" else "NONE"
        descent_count = int(descending)

        tel_line = (
            f"TEL,{seq},{timestamp},{temp:.2f},{pressure:.2f},{alt:.2f},"
            f"{ax:.3f},{ay:.3f},{az:.3f},{gx:.2f},{gy:.2f},{gz:.2f},"
            f"{self._test_peak_alt:.2f},{descent_count},{deployed},{deploy_reason}"
        )
        self.handle_line(tel_line)

        # Moving GPS track around the ground station. Longitude is corrected by
        # latitude so the path is roughly circular on the map.
        base_lat = self._gs_lat_default
        base_lon = self._gs_lon_default
        radius_deg = 0.0030
        angle = t / 18.0
        lat = base_lat + radius_deg * math.cos(angle)
        lon = base_lon + (radius_deg * math.sin(angle)) / max(0.2, math.cos(math.radians(base_lat)))
        sats = 9 + int(3.0 + 3.0 * math.sin(t / 9.0))
        gps_line = f"GPS,{seq},{timestamp},{lat:.7f},{lon:.7f},{alt:.2f},{sats},1"
        self.handle_line(gps_line)

        if int(t) - self._test_preview_last >= 10:
            self._test_preview_last = int(t)
            self._make_test_preview(seq, alt, temp, lat, lon)

        self._test_timer_id = self.root.after(1000, self._run_test_tick)

    def _make_test_preview(self, seq, alt, temp, lat, lon):
        """Generate a small fake camera preview frame for dashboard testing."""
        try:
            from PIL import ImageDraw
            img = Image.new("RGB", (640, 360), "#10151F")
            draw = ImageDraw.Draw(img)
            horizon = 165 + int(28 * math.sin(seq / 5.0))
            draw.rectangle((0, 0, 640, horizon), fill="#274D83")
            draw.rectangle((0, horizon, 640, 360), fill="#4C3A25")
            draw.line((0, horizon, 640, horizon), fill="#E0D25B", width=4)
            draw.ellipse((288, 138, 352, 202), outline="#00FFE5", width=3)
            draw.line((250, 170, 390, 170), fill="#00FFE5", width=2)
            draw.text((18, 18), "TEST PREVIEW", fill="#00FFE5")
            draw.text((18, 44), f"SEQ {seq:04d}", fill="white")
            draw.text((18, 70), f"ALT {alt:.1f} m", fill="white")
            draw.text((18, 96), f"TEMP {temp:.1f} C", fill="white")
            draw.text((18, 122), f"GPS {lat:.5f}, {lon:.5f}", fill="white")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            self.show_preview(buf.getvalue())
        except Exception as exc:
            self.log(f"Test preview generation failed: {exc}")

    def handle_line(self, line):
        with open(self.raw_log, "a", encoding="utf-8") as f:
            f.write(line + "\n")

        if line.startswith("TEL,"):
            self.handle_telemetry(line)
        elif line.startswith("GPS,"):
            self.handle_gps(line)
        elif line.startswith("STATUS,"):
            self.handle_status(line)
        elif line.startswith("IMGMETA,"):
            self.handle_imgmeta(line)
        elif line.startswith("IMG,"):
            self.handle_imgchunk(line)
        else:
            self.log(f"RAW {line}")

    def _safe_float(self, value, default=0.0):
        try:
            if str(value).lower() == "nan":
                return default
            return float(value)
        except Exception:
            return default

    def accel_to_attitude(self, ax, ay, az):
        roll = math.degrees(math.atan2(ay, az if abs(az) > 1e-6 else 1e-6))
        pitch = math.degrees(math.atan2(-ax, math.sqrt(ay * ay + az * az) + 1e-6))
        return roll, pitch


    def _init_mission_state(self):
        """Load/create persistent mission and row counters.

        row_id is a ground-station-side counter. It keeps increasing even when
        the CanSat seq counter restarts or the simulator starts from zero.
        mission_id identifies the app logging session/mission in every CSV and
        Google Sheets row.
        """
        self.mission_state_path = os.path.join(self.secrets_dir, "mission_state.json")
        now_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.mission_id = now_id
        self._mission_counters = {"telemetry": 0, "gps": 0}

        try:
            if os.path.isfile(self.mission_state_path):
                with open(self.mission_state_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.mission_id = str(data.get("mission_id") or now_id)
                counters = data.get("counters", {})
                self._mission_counters["telemetry"] = int(counters.get("telemetry", 0))
                self._mission_counters["gps"] = int(counters.get("gps", 0))
        except Exception:
            self.mission_id = now_id
            self._mission_counters = {"telemetry": 0, "gps": 0}

        self._ensure_csv_schema(self.telemetry_csv, self._telemetry_csv_headers(), "telemetry")
        self._ensure_csv_schema(self.gps_csv, self._gps_csv_headers(), "gps")
        self._ensure_csv_schema(self.telemetry_test_csv, self._telemetry_csv_headers(), "telemetry")
        self._ensure_csv_schema(self.gps_test_csv, self._gps_csv_headers(), "gps")
        self._sync_counters_from_csv()
        self._save_mission_state()

    def _telemetry_csv_headers(self):
        return [
            "row_id", "mission_id", "source",
            "seq", "timestamp", "temp_C", "pressure_hPa", "alt_m",
            "ax", "ay", "az", "gx", "gy", "gz",
            "peak_alt_m", "descent_count", "deployed", "deploy_reason",
        ]

    def _gps_csv_headers(self):
        return [
            "row_id", "mission_id", "source",
            "seq", "timestamp", "latitude", "longitude",
            "gnss_alt_m", "satellites", "fix_quality",
        ]

    def _current_row_source(self):
        return "TEST" if getattr(self, "_test_mode_running", False) else "REAL"

    def _active_telemetry_csv(self):
        return self.telemetry_test_csv if self._current_row_source() == "TEST" else self.telemetry_csv

    def _active_gps_csv(self):
        return self.gps_test_csv if self._current_row_source() == "TEST" else self.gps_csv

    def _notify_csv_written(self, source, csv_path):
        if source == "TEST":
            self.log(f"TEST row saved to {os.path.basename(csv_path)}")

    def _ensure_csv_schema(self, path, expected_headers, counter_key):
        """Upgrade old CSV files by adding row_id and mission_id columns.

        Old files are not deleted. A .bak copy is made before rewriting.
        """
        if not os.path.exists(path):
            return
        try:
            with open(path, newline="", encoding="utf-8") as f:
                rows = list(csv.reader(f))
            if not rows:
                return
            current_headers = rows[0]
            if current_headers == expected_headers:
                return

            backup = path + ".pre_row_id.bak"
            if not os.path.exists(backup):
                with open(backup, "w", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerows(rows)

            migrated = [expected_headers]
            start = int(self._mission_counters.get(counter_key, 0))
            has_row_id = current_headers[:2] == ["row_id", "mission_id"]
            has_source = "source" in current_headers
            for i, row in enumerate(rows[1:], start=1):
                if has_row_id:
                    new_row = list(row)
                    if not has_source:
                        # Existing mission rows were collected before TEST/REAL tagging existed.
                        # Treat them as REAL instead of guessing.
                        new_row = new_row[:2] + ["REAL"] + new_row[2:]
                    migrated.append(new_row[:len(expected_headers)] + [""] * max(0, len(expected_headers) - len(new_row)))
                    continue
                migrated.append([start + i, self.mission_id, "REAL"] + row)
            with open(path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerows(migrated)
            self._mission_counters[counter_key] = max(start, start + max(0, len(rows) - 1))
        except Exception as exc:
            try:
                self.log(f"CSV schema upgrade skipped for {os.path.basename(path)}: {exc}")
            except Exception:
                pass

    def _sync_counters_from_csv(self):
        csv_paths = [
            ("telemetry", self.telemetry_csv),
            ("telemetry", self.telemetry_test_csv),
            ("gps", self.gps_csv),
            ("gps", self.gps_test_csv),
        ]
        for key, path in csv_paths:
            max_id = 0
            if os.path.exists(path):
                try:
                    with open(path, newline="", encoding="utf-8") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            try:
                                max_id = max(max_id, int(row.get("row_id") or 0))
                            except Exception:
                                pass
                except Exception:
                    pass
            self._mission_counters[key] = max(int(self._mission_counters.get(key, 0)), max_id)

    def _save_mission_state(self):
        try:
            os.makedirs(self.secrets_dir, exist_ok=True)
            data = {
                "mission_id": self.mission_id,
                "counters": self._mission_counters,
                "updated_at": datetime.now().isoformat(timespec="seconds"),
            }
            tmp = self.mission_state_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            os.replace(tmp, self.mission_state_path)
        except Exception as exc:
            try:
                self.log(f"Mission state save failed: {exc}")
            except Exception:
                pass

    def _next_row_id(self, counter_key):
        self._mission_counters[counter_key] = int(self._mission_counters.get(counter_key, 0)) + 1
        self._save_mission_state()
        return self._mission_counters[counter_key]

    def handle_gps(self, line):
        parts = line.split(",")

        if len(parts) < 8:
            self.log(f"Bad GPS packet: {line}")
            return

        try:
            _, seq, timestamp, lat, lon, gnss_alt, sats, fix = parts[:8]
        except ValueError:
            self.log(f"Parse GPS failed: {line}")
            return

        self.card_vars["gps_lat"].set(lat)
        self.card_vars["gps_lon"].set(lon)
        self.card_vars["gps_alt"].set(gnss_alt)
        self.card_vars["gps_sats"].set(sats)

        self.gps_packet_count += 1

        # ── Update live map ──────────────────────────────────────────
        try:
            lat_f = float(lat)
            lon_f = float(lon)
            alt_f = float(gnss_alt) if gnss_alt else 0.0
            if lat_f != 0.0 and lon_f != 0.0:
                self.current_gps_lat = lat_f
                self.current_gps_lon = lon_f
                self.current_gps_alt = alt_f
                self._collect_ground_station_average(lat_f, lon_f)
                if self._live_map:
                    self._live_map.update_cansat(lat_f, lon_f, alt_f)
                if self._google_earth:
                    self._google_earth.update_position(lat_f, lon_f, alt_f, timestamp=timestamp)
        except (ValueError, TypeError):
            pass

        row_id = self._next_row_id("gps")
        source = self._current_row_source()
        gps_csv_path = self._active_gps_csv()
        new_file = not os.path.exists(gps_csv_path)
        with open(gps_csv_path, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(self._gps_csv_headers())
            w.writerow([row_id, self.mission_id, source, seq, timestamp, lat, lon, gnss_alt, sats, fix])
        self._notify_csv_written(source, gps_csv_path)

        # Notify the Google Sheets uploader that new data is available
        if self._sheets_uploader:
            self._sheets_uploader.notify()

        self.log(
            f"GPS seq={seq} | lat={lat} | lon={lon} | "
            f"alt={gnss_alt} m | sats={sats} | fix={fix}"
        )

    def handle_telemetry(self, line):
        parts = line.split(",")

        if len(parts) < 12:
            self.log(f"Bad TEL packet: {line}")
            return

        try:
            if len(parts) >= 16:
                (
                    _, seq, timestamp, temp, pressure, alt,
                    ax, ay, az, gx, gy, gz,
                    peak_alt, descent_count, deployed, deploy_reason
                ) = parts[:16]
            else:
                (
                    _, seq, timestamp, temp, pressure, alt,
                    ax, ay, az, gx, gy, gz
                ) = parts[:12]
                peak_alt = alt
                descent_count = "0"
                deployed = "0"
                deploy_reason = "NONE"
        except ValueError:
            self.log(f"Parse TEL failed: {line}")
            return

        try:
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.write(f"ACK,{seq}\n".encode("utf-8"))
        except Exception as e:
            self.log(f"ACK send error: {e}")

        self.packet_count += 1
        self.last_packet_time = datetime.now()
        self.packet_var.set(f"Packets: {self.packet_count}")
        self.last_rx_var.set(f"Last RX: {self.last_packet_time.strftime('%H:%M:%S')}")

        try:
            seq_int = int(seq)
            if self.last_seq is not None and seq_int != self.last_seq + 1:
                self.log(f"PACKET LOSS: expected {self.last_seq + 1}, got {seq_int}")
            self.last_seq = seq_int
        except Exception:
            pass

        self.card_vars["seq"].set(seq)
        self.card_vars["temp"].set(temp)
        self.card_vars["pressure"].set(pressure)
        self.card_vars["alt"].set(alt)
        self.card_vars["peak_alt"].set(peak_alt)
        self.card_vars["deploy"].set(f"{deployed} {deploy_reason}")
        self.card_vars["ax"].set(ax)
        self.card_vars["ay"].set(ay)
        self.card_vars["az"].set(az)
        self.card_vars["gx"].set(gx)
        self.card_vars["gy"].set(gy)
        self.card_vars["gz"].set(gz)
        self.timestamp_var.set(timestamp)

        axf = self._safe_float(ax)
        ayf = self._safe_float(ay)
        azf = self._safe_float(az, default=1.0)
        gxf = self._safe_float(gx)
        gyf = self._safe_float(gy)

        roll_new, pitch_new = self.accel_to_attitude(axf, ayf, azf)

        alpha_roll = 0.12
        alpha_pitch = 0.12
        self.roll = (1 - alpha_roll) * self.roll + alpha_roll * roll_new
        self.pitch = (1 - alpha_pitch) * self.pitch + alpha_pitch * pitch_new

        self.roll = max(-85, min(85, self.roll))
        self.pitch = max(-40, min(40, self.pitch))

        self.vx = max(-60, min(60, gyf * 0.08))
        self.vy = max(-60, min(60, -gxf * 0.08))

        self.attitude.draw_indicator(self.roll, self.pitch)

        self.telemetry_history.append({
            "time": datetime.now(),
            "temp": self._safe_float(temp),
            "pressure": self._safe_float(pressure),
            "alt": self._safe_float(alt),
        })
        if len(self.telemetry_history) > 2000:
            cutoff = datetime.now() - timedelta(seconds=600)
            self.telemetry_history = [d for d in self.telemetry_history if d["time"] > cutoff]

        row_id = self._next_row_id("telemetry")
        source = self._current_row_source()
        telemetry_csv_path = self._active_telemetry_csv()
        new_file = not os.path.exists(telemetry_csv_path)
        with open(telemetry_csv_path, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(self._telemetry_csv_headers())
            w.writerow([
                row_id, self.mission_id, source,
                seq, timestamp, temp, pressure, alt,
                ax, ay, az, gx, gy, gz,
                peak_alt, descent_count, deployed, deploy_reason
            ])
        self._notify_csv_written(source, telemetry_csv_path)

        # Notify the Google Sheets uploader that new data is available
        if self._sheets_uploader:
            self._sheets_uploader.notify()

    def handle_imgmeta(self, line):
        parts = line.split(",", 4)
        if len(parts) != 5:
            self.log(f"Bad IMGMETA packet: {line}")
            return

        _, image_seq, timestamp, filename, total = parts
        try:
            total = int(total)
        except ValueError:
            self.log(f"Bad IMGMETA total: {line}")
            return

        self.image_buffers[image_seq] = {
            "timestamp": timestamp,
            "filename": filename,
            "total": total,
            "chunks": {}
        }
        self.log(f"IMGMETA image_seq={image_seq} chunks={total}")

    def handle_imgchunk(self, line):
        parts = line.split(",", 4)
        if len(parts) != 5:
            self.log(f"Bad IMG packet: {line}")
            return

        _, image_seq, chunk_index, total, chunk_data = parts

        try:
            chunk_index = int(chunk_index)
            total = int(total)
        except ValueError:
            self.log(f"Bad IMG numbering: {line}")
            return

        if image_seq not in self.image_buffers:
            self.image_buffers[image_seq] = {
                "timestamp": "",
                "filename": f"preview_{image_seq}.jpg",
                "total": total,
                "chunks": {}
            }

        buf = self.image_buffers[image_seq]
        buf["total"] = total
        buf["chunks"][chunk_index] = chunk_data

        if len(buf["chunks"]) == total:
            self.reassemble_image(image_seq)

    def reassemble_image(self, image_seq):
        buf = self.image_buffers.get(image_seq)
        if not buf:
            return

        total = buf["total"]
        chunks = buf["chunks"]

        if any(i not in chunks for i in range(total)):
            self.log(f"IMG image_seq={image_seq} incomplete")
            return

        try:
            b64_data = "".join(chunks[i] for i in range(total))
            img_bytes = base64.b64decode(b64_data)

            safe_name = os.path.splitext(buf["filename"])[0] + f"_preview_{image_seq}.jpg"
            out_path = os.path.join(self.preview_dir, safe_name)

            with open(out_path, "wb") as f:
                f.write(img_bytes)

            self.show_preview(img_bytes)
            self.log(f"Preview saved: {out_path}")

            del self.image_buffers[image_seq]

        except Exception as e:
            self.log(f"Preview decode failed: {e}")

    def show_preview(self, img_bytes):
        try:
            image = Image.open(io.BytesIO(img_bytes))
            image.thumbnail((420, 320))
            photo = ImageTk.PhotoImage(image)
            self.current_preview_photo = photo
            self.preview_label.configure(image=photo, text="")
        except Exception as e:
            self.log(f"Preview display failed: {e}")

    def _open_map_window(self):
        """Open (or raise) the live satellite map window."""
        ensure_map_dependencies(parent=self.root, log_fn=self.log, auto_install=True)
        if self._tile_cache is None:
            self._tile_cache = LocalEsriTileCache(self.tile_cache_dir, log_fn=self.log)
            self._tile_cache.start()
        self._live_map._tile_cache = self._tile_cache
        self._live_map.show()

    def _open_google_earth(self):
        """Open the live KML loader in Google Earth Pro or the default KML app."""
        if self._google_earth is None:
            self._google_earth = GoogleEarthLink(self.kml_dir, log_fn=self.log)
            self._google_earth.update_ground(self._gs_lat_default, self._gs_lon_default)
        ok, msg = self._google_earth.open()
        self.log(msg)
        if not ok:
            messagebox.showerror("Google Earth", msg)

    def _resolve_config_path(self, path_value):
        """Resolve user-entered paths; relative paths are relative to the script."""
        if not path_value:
            return ""
        path_value = os.path.expandvars(os.path.expanduser(path_value.strip()))
        if os.path.isabs(path_value):
            return path_value
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.abspath(os.path.join(base_dir, path_value))

    def _get_sheets_config(self):
        """
        Return (sheet_url, creds_path, source_name) or None.

        Priority:
          1. Environment variables
          2. Built-in constants near the top of this file
          3. Saved config in CanSat_GroundStation_Data/Secrets/sheets_config.json
        """
        env_url = os.environ.get(SHEETS_URL_ENV, "").strip()
        env_creds = os.environ.get(SHEETS_CREDS_ENV, "").strip()
        if env_url and env_creds:
            return env_url, self._resolve_config_path(env_creds), "environment variables"

        if DEFAULT_GOOGLE_SHEET_URL.strip() and DEFAULT_SERVICE_ACCOUNT_JSON.strip():
            return (
                DEFAULT_GOOGLE_SHEET_URL.strip(),
                self._resolve_config_path(DEFAULT_SERVICE_ACCOUNT_JSON),
                "built-in code settings",
            )

        try:
            if os.path.isfile(self.sheets_config_path):
                with open(self.sheets_config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                url = str(data.get("sheet_url", "")).strip()
                creds = str(data.get("service_account_json", "")).strip()
                if url and creds:
                    return url, self._resolve_config_path(creds), "saved settings"
        except Exception as exc:
            self.log(f"Could not read saved Sheets config: {exc}")

        return None

    def _save_sheets_config(self, sheet_url, creds_path):
        """Save only the sheet URL and JSON file path, not the private key contents."""
        try:
            os.makedirs(self.secrets_dir, exist_ok=True)
            data = {
                "sheet_url": sheet_url,
                "service_account_json": creds_path,
            }
            with open(self.sheets_config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.log(f"Saved Sheets config: {self.sheets_config_path}")
        except Exception as exc:
            self.log(f"Could not save Sheets config: {exc}")

    def toggle_sheets_upload(self):
        """Start or stop the Google Sheets live uploader."""
        if self._sheets_uploader is not None:
            # Already running → stop it
            self._sheets_uploader.stop()
            self._sheets_uploader = None
            self._sheets_btn.config(text="☁  SHEETS", bg="#1A237E")
            self.log("Google Sheets upload stopped.")
            return

        if not ensure_sheets_dependencies(
            parent=self.root,
            log_fn=self.log,
            auto_install=True,
        ):
            messagebox.showerror(
                "Google Sheets dependency install failed",
                sheets_dependency_message(),
            )
            return

        config = self._get_sheets_config()
        if config is None:
            dlg = SheetsConfigDialog(self.root)
            self.root.wait_window(dlg)

            if dlg.result is None:
                return   # user cancelled

            sheet_url, creds_path, save_config = dlg.result
            creds_path = self._resolve_config_path(creds_path)
            if save_config:
                self._save_sheets_config(sheet_url, creds_path)
        else:
            sheet_url, creds_path, source_name = config
            self.log(f"Using Google Sheets config from {source_name}.")

        if not os.path.isfile(creds_path):
            messagebox.showerror(
                "Error",
                "Service Account JSON file not found:\n"
                f"{creds_path}\n\n"
                "Fix DEFAULT_SERVICE_ACCOUNT_JSON near the top of the code, "
                "or edit/delete the saved sheets_config.json file.",
            )
            return

        self._sheets_uploader = GoogleSheetsUploader(
            sheet_url=sheet_url,
            credentials=creds_path,
            telemetry_csv=self.telemetry_csv,
            gps_csv=self.gps_csv,
            telemetry_test_csv=self.telemetry_test_csv,
            gps_test_csv=self.gps_test_csv,
            log_fn=self.log,
        )
        self._sheets_uploader.start()
        self._sheets_btn.config(text="☁  SHEETS ●", bg="#0D47A1")
        self.log("Google Sheets upload started.")

    def shutdown(self):
        if getattr(self, "_test_mode_running", False):
            self.stop_test_mode()
        if self._sheets_uploader:
            self._sheets_uploader.stop()
        if self._tile_cache:
            self._tile_cache.stop()
        self.disconnect()
        self.root.destroy()


# ---------------------------------------------------------------------------
# Integrated fullscreen dashboard classes
# ---------------------------------------------------------------------------

class EmbeddedTelemetryViz(TelemetryVizWindow):
    """Telemetry charts embedded in the main dashboard instead of a Toplevel."""
    def __init__(self, app, parent):
        self.app = app
        self.win = parent
        self.time_window_sec = tk.IntVar(value=60)
        self._att_var = tk.StringVar(value="ROLL  +0.0 deg   PITCH  +0.0 deg")
        self._build_ui()
        self._schedule_chart_update()

    def _on_close(self):
        return

    def _build_ui(self):
        chart_parent = ttk.Frame(self.win, padding=(6, 4, 6, 4))
        chart_parent.pack(fill="both", expand=True)
        self._build_charts(chart_parent)


class Embedded3DOrientation(TelemetryVizWindow):
    """3-D satellite orientation view embedded as a dashboard tab."""
    def __init__(self, app, parent):
        self.app = app
        self.win = parent
        self.time_window_sec = tk.IntVar(value=60)
        self._att_var = tk.StringVar(value="ROLL  +0.0 deg   PITCH  +0.0 deg")
        self._build_ui()
        self._schedule_orientation_update()

    def _on_close(self):
        return

    def _build_ui(self):
        frame = ttk.Frame(self.win, padding=(6, 4, 6, 4))
        frame.pack(fill="both", expand=True)
        self._build_3d(frame)

    def _schedule_orientation_update(self):
        if not self.win.winfo_exists():
            return
        if hasattr(self, "_gl") and not self._gl.imu_driven:
            self._att_var.set(
                f"AZ  {self._gl._mouse_ry:+.1f} deg   EL  {self._gl._mouse_rx:+.1f} deg  [mouse]"
            )
        else:
            self._att_var.set(
                f"ROLL  {self.app.roll:+.1f} deg   PITCH  {self.app.pitch:+.1f} deg"
            )
        self.win.after(250, self._schedule_orientation_update)


class EmbeddedLiveMap(LiveMapWindow):
    """Live map embedded in the main dashboard instead of a separate window."""
    def __init__(self, root, app, parent):
        self._embedded_parent = parent
        super().__init__(root, app)
        self._create_window()

    def show(self):
        self._tile_cache = getattr(self.app, "_tile_cache", self._tile_cache)
        if _HAS_MAPVIEW and getattr(self, "_map", None) is not None:
            self._apply_tile_server()
            if self._cansat_lat is not None:
                self._refresh_map()
            elif self._gs_lat is not None:
                self._map.set_position(self._gs_lat, self._gs_lon)
        self._maybe_prefetch_esri_tiles()

    def _create_window(self):
        self._win = self._embedded_parent
        for child in self._win.winfo_children():
            child.destroy()

        self._dist_var = tk.StringVar(value="-")
        self._bearing_var = tk.StringVar(value="-")
        self._csat_ll_var = tk.StringVar(value="-")
        self._gs_ll_var = tk.StringVar(value="set below")
        self._tile_var = tk.StringVar(value="Esri Satellite")
        self._tile_cache_var = tk.StringVar(value="cache: ready")

        self._build_ui()

    def _build_ui(self):
        win = self._win

        title_bar = tk.Frame(win, bg="#05080F", pady=3)
        title_bar.pack(fill="x")
        tk.Label(
            title_bar,
            text="CANSAT LIVE SATELLITE MAP",
            bg="#05080F",
            fg=self._CYAN,
            font=("Consolas", 11, "bold"),
        ).pack(side="left", padx=8)

        info = tk.Frame(win, bg="#0A0E14", pady=2)
        info.pack(fill="x")
        for i in range(8):
            info.columnconfigure(i, weight=1 if i in (1, 3, 5, 7) else 0)

        def _lbl(text, col):
            tk.Label(
                info,
                text=text,
                bg="#0A0E14",
                fg=self._DIM,
                font=("Consolas", 8),
            ).grid(row=0, column=col, sticky="w", padx=(6, 1))

        def _val(var, col, color=None, width=16):
            tk.Label(
                info,
                textvariable=var,
                bg="#0A0E14",
                fg=color or self._CYAN,
                font=("Consolas", 9, "bold"),
                width=width,
                anchor="w",
            ).grid(row=0, column=col, sticky="ew", padx=1)

        _lbl("DIST", 0); _val(self._dist_var, 1, width=12)
        _lbl("BRG", 2); _val(self._bearing_var, 3, width=12)
        _lbl("SAT", 4); _val(self._csat_ll_var, 5, width=17)
        _lbl("GND", 6); _val(self._gs_ll_var, 7, color=self._MAGENTA, width=17)

        controls = tk.Frame(win, bg="#0A0E14", pady=3)
        controls.pack(fill="x")
        for i in range(12):
            controls.columnconfigure(i, weight=0)
        controls.columnconfigure(11, weight=1)

        tk.Label(controls, text="GND LAT", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).grid(row=0, column=0, sticky="w", padx=(6, 2))
        self._gs_lat_entry = tk.Entry(
            controls, width=11, font=("Consolas", 8), bg="#11151C", fg=self._MAGENTA,
            insertbackground=self._MAGENTA, relief="flat",
        )
        self._gs_lat_entry.grid(row=0, column=1, sticky="w", padx=(0, 5))

        tk.Label(controls, text="LON", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).grid(row=0, column=2, sticky="w", padx=(0, 2))
        self._gs_lon_entry = tk.Entry(
            controls, width=11, font=("Consolas", 8), bg="#11151C", fg=self._MAGENTA,
            insertbackground=self._MAGENTA, relief="flat",
        )
        self._gs_lon_entry.grid(row=0, column=3, sticky="w", padx=(0, 5))

        tk.Button(
            controls, text="SET", font=("Consolas", 8, "bold"), bg=self._MAGENTA,
            fg="#000", activebackground="#FF33AA", relief="flat", padx=8,
            command=self._on_set_gs,
        ).grid(row=0, column=4, sticky="w", padx=(0, 6))

        tk.Button(
            controls, text="GND=SAT", font=("Consolas", 8, "bold"), bg="#7B2CBF",
            fg="white", activebackground="#9D4EDD", relief="flat", padx=7,
            command=self.app.set_ground_station_from_current_gps,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=(6, 5), pady=(3, 0))

        tk.Button(
            controls, text="AVG GND", font=("Consolas", 8, "bold"), bg="#5A189A",
            fg="white", activebackground="#7B2CBF", relief="flat", padx=7,
            command=lambda: self.app.start_ground_station_average(20),
        ).grid(row=1, column=2, columnspan=2, sticky="w", padx=(0, 5), pady=(3, 0))

        tk.Button(
            controls, text="PC LOC", font=("Consolas", 8, "bold"), bg="#006D77",
            fg="white", activebackground="#118AB2", relief="flat", padx=7,
            command=self.app.set_ground_station_from_pc_location,
        ).grid(row=1, column=4, columnspan=2, sticky="w", padx=(0, 5), pady=(3, 0))

        tk.Label(controls, text="MAP", bg="#0A0E14", fg=self._DIM,
                 font=("Consolas", 8)).grid(row=0, column=5, sticky="w", padx=(0, 2))
        self._tile_combo = ttk.Combobox(
            controls,
            textvariable=self._tile_var,
            values=list(self._TILE_SERVERS.keys()),
            width=15,
            state="readonly",
        )
        self._tile_combo.grid(row=0, column=6, sticky="w", padx=(0, 6))
        self._tile_combo.bind("<<ComboboxSelected>>", self._on_tile_change)

        tk.Button(
            controls, text="PREFETCH", font=("Consolas", 8, "bold"), bg="#264653",
            fg="white", activebackground="#2A9D8F", relief="flat", padx=7,
            command=self._prefetch_esri_tiles,
        ).grid(row=0, column=7, sticky="w", padx=(0, 5))

        tk.Label(
            controls, textvariable=self._tile_cache_var, bg="#0A0E14", fg=self._DIM,
            font=("Consolas", 8), width=15, anchor="w",
        ).grid(row=0, column=8, sticky="w", padx=(0, 4))

        tk.Button(
            controls, text="GOOGLE MAPS", font=("Consolas", 8, "bold"), bg="#2D6A4F",
            fg="white", activebackground="#40916C", relief="flat", padx=7,
            command=self._open_google_maps_satellite,
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=(6, 5), pady=(3, 0))

        tk.Button(
            controls, text="GOOGLE EARTH", font=("Consolas", 8, "bold"), bg="#3A0CA3",
            fg="white", activebackground="#5F37D6", relief="flat", padx=7,
            command=self._open_google_earth,
        ).grid(row=2, column=2, columnspan=2, sticky="w", padx=(0, 5), pady=(3, 0))

        tk.Button(
            controls, text="GE PATH", font=("Consolas", 8, "bold"), bg="#222831",
            fg="white", activebackground="#393E46", relief="flat", padx=7,
            command=self._choose_google_earth_path,
        ).grid(row=2, column=4, columnspan=2, sticky="w", padx=(0, 5), pady=(3, 0))

        map_frame = tk.Frame(win, bg=self._BG)
        map_frame.pack(fill="both", expand=True)

        if _HAS_MAPVIEW:
            self._map = tkintermapview.TkinterMapView(map_frame, corner_radius=0)
            self._map.pack(fill="both", expand=True)
            self._apply_tile_server()
            self._map.set_zoom(17)
        else:
            self._map = None
            self._canvas = tk.Canvas(map_frame, bg=self._BG, highlightthickness=0)
            self._canvas.pack(fill="both", expand=True)
            self._canvas.bind("<Configure>", lambda _e: self._redraw_fallback())
            self._root.after(120, self._redraw_fallback)


class FullscreenGroundStationApp(GroundStationApp):
    """One-window fullscreen dashboard layout."""
    def __init__(self, root):
        self.root = root
        self.root.title("CanSat Ground Station - Full Dashboard")
        self.root.configure(bg="#0F1115")
        self._fullscreen = False

        base_dir = os.path.dirname(os.path.abspath(__file__))

        self.output_root = os.path.join(base_dir, "CanSat_GroundStation_Data")
        self.csv_dir = os.path.join(self.output_root, "CSVs")
        self.kml_dir = os.path.join(self.output_root, "KMLs")
        self.preview_dir = os.path.join(self.output_root, "Previews")
        self.raw_log_dir = os.path.join(self.output_root, "Raw_Logs")
        self.tile_cache_dir = os.path.join(self.output_root, "Tile_Cache")
        self.secrets_dir = os.path.join(self.output_root, "Secrets")
        self.sheets_config_path = os.path.join(self.secrets_dir, "sheets_config.json")

        for folder in (
            self.output_root,
            self.csv_dir,
            self.kml_dir,
            self.preview_dir,
            self.raw_log_dir,
            self.tile_cache_dir,
            self.secrets_dir,
        ):
            os.makedirs(folder, exist_ok=True)

        self.log_dir = self.output_root
        self.telemetry_csv = os.path.join(self.csv_dir, "telemetry.csv")
        self.gps_csv = os.path.join(self.csv_dir, "gps.csv")
        self.telemetry_test_csv = os.path.join(self.csv_dir, "telemetry_TEST.csv")
        self.gps_test_csv = os.path.join(self.csv_dir, "gps_TEST.csv")
        self.raw_log = os.path.join(self.raw_log_dir, "raw_packets.log")
        self._init_mission_state()

        self.serial_port = None
        self.serial_thread = None
        self.running = False
        self.rx_queue = queue.Queue()
        self.sat_started = False

        self.packet_count = 0
        self.gps_packet_count = 0
        self.last_packet_time = None
        self.current_preview_photo = None
        self.image_buffers = {}

        self.roll = 0.0
        self.pitch = 0.0
        self.vx = 0.0
        self.vy = 0.0

        self.last_seq = None
        self.telemetry_history = []

        self._sheets_uploader = None
        self._live_map = None
        self._google_earth = GoogleEarthLink(self.kml_dir)
        self._tile_cache = LocalEsriTileCache(self.tile_cache_dir)
        self._tile_cache.start()

        self.ground_station_config_path = os.path.join(self.secrets_dir, "ground_station_location.json")
        self._gs_lat_default = 59.4370
        self._gs_lon_default = 24.7536
        self._load_ground_station_location()
        self.current_gps_lat = None
        self.current_gps_lon = None
        self.current_gps_alt = None
        self._gnd_average_points = []
        self._gnd_average_remaining = 0

        self._test_mode_running = False
        self._test_seq = 0
        self._test_start_time = None
        self._test_timer_id = None
        self._test_peak_alt = 0.0
        self._test_preview_last = -999

        self._build_styles()
        ensure_map_dependencies(parent=self.root, log_fn=lambda msg: print("[Map] " + msg), auto_install=True)
        self._build_ui()

        self._tile_cache.log = self.log
        self._google_earth.log = self.log
        self._google_earth.update_ground(self._gs_lat_default, self._gs_lon_default)
        if self._live_map:
            self._live_map._tile_cache = self._tile_cache
            self._live_map._apply_tile_server()
            self._live_map.set_ground_station(self._gs_lat_default, self._gs_lon_default)

        self.log(f"Output folder: {self.output_root}")
        self.log(f"CSV folder: {self.csv_dir}")
        self.log(f"Real CSVs: telemetry.csv, gps.csv")
        self.log(f"Test CSVs: telemetry_TEST.csv, gps_TEST.csv")
        self.log(f"KML folder: {self.kml_dir}")
        self.log(f"Tile cache folder: {self.tile_cache_dir}")
        self.log(f"Sheets config: {self.sheets_config_path}")
        self.log(f"Mission ID: {self.mission_id}")

        self.refresh_ports()
        self.root.after(100, self.process_queue)
        self.root.after(500, self.update_link_status)
        self._activate_dashboard_window_mode()

    def _load_ground_station_location(self):
        """Load saved ground-station coordinates if available."""
        try:
            if not os.path.exists(self.ground_station_config_path):
                return
            with open(self.ground_station_config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            lat = float(data.get("lat"))
            lon = float(data.get("lon"))
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                self._gs_lat_default = lat
                self._gs_lon_default = lon
        except Exception:
            pass

    def _save_ground_station_location(self):
        """Persist ground-station coordinates for the next launch."""
        try:
            os.makedirs(self.secrets_dir, exist_ok=True)
            with open(self.ground_station_config_path, "w", encoding="utf-8") as f:
                json.dump({
                    "lat": self._gs_lat_default,
                    "lon": self._gs_lon_default,
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                }, f, indent=2)
        except Exception as exc:
            self.log(f"Ground-station save failed: {exc}")

    def set_ground_station_location(self, lat, lon, source="manual"):
        """Set GND position everywhere: app state, map, Google Earth and saved config."""
        try:
            lat = float(lat)
            lon = float(lon)
        except Exception:
            messagebox.showerror("Ground station", "Invalid ground-station coordinates.")
            return False
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            messagebox.showerror("Ground station", "Latitude/longitude are out of range.")
            return False

        self._gs_lat_default = lat
        self._gs_lon_default = lon
        self._save_ground_station_location()

        if self._google_earth:
            self._google_earth.update_ground(lat, lon)
        if self._live_map:
            self._live_map.set_ground_station(lat, lon)
            try:
                self._live_map._prefetch_esri_tiles()
            except Exception:
                pass
        self.log(f"Ground station set from {source}: {lat:.7f}, {lon:.7f}")
        return True

    def set_ground_station_from_current_gps(self):
        """Use the latest received GPS coordinate as the ground station."""
        if self.current_gps_lat is None or self.current_gps_lon is None:
            messagebox.showinfo(
                "Ground station",
                "No valid GPS position has been received yet. Wait for a GPS fix, then press this again.",
            )
            return
        self.set_ground_station_location(self.current_gps_lat, self.current_gps_lon, source="latest GPS fix")

    def set_ground_station_from_pc_location(self):
        """Use the Windows location service as the ground-station position."""
        if platform.system().lower() != "windows":
            messagebox.showinfo(
                "Ground station",
                "PC location is only available on Windows. Use GND=SAT, AVG GND, or manual coordinates on this system.",
            )
            return

        self.log("Requesting PC location from Windows Location Service...")
        thread = threading.Thread(target=self._pc_location_worker, daemon=True, name="PcLocationWorker")
        thread.start()

    def _pc_location_worker(self):
        try:
            lat, lon, acc = self._query_windows_location()
            self.root.after(0, lambda: self._apply_pc_location(lat, lon, acc))
        except Exception as exc:
            self.root.after(0, lambda exc=exc: self._pc_location_failed(str(exc)))

    def _pc_location_failed(self, msg):
        self.log(f"PC location failed: {msg}")
        messagebox.showerror(
            "PC location failed",
            "Could not read Windows Location Service position.\n\n"
            "Check Windows Settings > Privacy & security > Location, and make sure Location services "
            "and desktop app access are enabled.\n\n"
            f"Details: {msg}",
        )

    def _apply_pc_location(self, lat, lon, accuracy_m):
        ok = self.set_ground_station_location(lat, lon, source=f"Windows PC location, accuracy ~{accuracy_m:.0f} m")
        if ok:
            messagebox.showinfo(
                "Ground station",
                f"Ground station set from Windows PC location.\n\n"
                f"Latitude: {lat:.7f}\nLongitude: {lon:.7f}\nAccuracy: about {accuracy_m:.0f} m",
            )

    def _query_windows_location(self):
        """Return (lat, lon, accuracy_m) from Windows Location Service via PowerShell."""
        powershell = shutil.which("powershell") or shutil.which("pwsh")
        if not powershell:
            raise RuntimeError("PowerShell was not found.")

        ps_script = r'''
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Device
$watcher = New-Object System.Device.Location.GeoCoordinateWatcher([System.Device.Location.GeoPositionAccuracy]::High)
$watcher.MovementThreshold = 1
$started = $watcher.TryStart($false, [TimeSpan]::FromSeconds(20))
if (-not $started) { throw 'Windows Location Service did not start or permission was denied.' }
$deadline = (Get-Date).AddSeconds(20)
do {
    Start-Sleep -Milliseconds 250
    $loc = $watcher.Position.Location
} while ($loc.IsUnknown -and (Get-Date) -lt $deadline)
if ($loc.IsUnknown) { throw 'Windows returned an unknown location. Wait for location lock or enable Location Services.' }
$lat = [double]$loc.Latitude
$lon = [double]$loc.Longitude
$acc = [double]$loc.HorizontalAccuracy
if ([double]::IsNaN($acc) -or $acc -le 0) { $acc = 0 }
Write-Output ("OK,{0:R},{1:R},{2:R}" -f $lat, $lon, $acc)
'''
        kwargs = {}
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(
            [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30,
            **kwargs,
        )
        if result.returncode != 0:
            details = (result.stderr or result.stdout or "PowerShell location query failed.").strip()
            raise RuntimeError(details[-1200:])
        line = ""
        for raw in result.stdout.splitlines():
            raw = raw.strip()
            if raw.startswith("OK,"):
                line = raw
                break
        if not line:
            raise RuntimeError((result.stdout or "No location result returned.").strip()[-1200:])
        parts = line.split(",")
        if len(parts) != 4:
            raise RuntimeError(f"Unexpected location result: {line}")
        lat = float(parts[1])
        lon = float(parts[2])
        acc = float(parts[3])
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            raise RuntimeError(f"Out-of-range location returned: {lat}, {lon}")
        return lat, lon, acc

    def start_ground_station_average(self, samples=20):
        """Average the next GPS fixes to reduce random GPS wander."""
        self._gnd_average_points = []
        self._gnd_average_remaining = int(samples)
        self.log(f"Ground-station averaging started: using next {samples} GPS fixes.")

    def _collect_ground_station_average(self, lat, lon):
        if self._gnd_average_remaining <= 0:
            return
        self._gnd_average_points.append((lat, lon))
        self._gnd_average_remaining -= 1
        self.log(f"GND averaging: {len(self._gnd_average_points)} sample(s), {self._gnd_average_remaining} left")
        if self._gnd_average_remaining == 0 and self._gnd_average_points:
            avg_lat = sum(p[0] for p in self._gnd_average_points) / len(self._gnd_average_points)
            avg_lon = sum(p[1] for p in self._gnd_average_points) / len(self._gnd_average_points)
            self.set_ground_station_location(avg_lat, avg_lon, source=f"{len(self._gnd_average_points)} GPS-fix average")

    def _activate_dashboard_window_mode(self):
        try:
            self.root.state("zoomed")
        except Exception:
            try:
                self.root.attributes("-zoomed", True)
            except Exception:
                w = self.root.winfo_screenwidth()
                h = self.root.winfo_screenheight()
                self.root.geometry(f"{w}x{h}+0+0")
        self.root.bind("<F11>", self._toggle_fullscreen)
        self.root.bind("<Escape>", self._leave_fullscreen)

    def _toggle_fullscreen(self, _event=None):
        self._fullscreen = not self._fullscreen
        self.root.attributes("-fullscreen", self._fullscreen)

    def _leave_fullscreen(self, _event=None):
        self._fullscreen = False
        self.root.attributes("-fullscreen", False)

    def _small_button(self, parent, text, command, bg, activebackground, padx=10):
        btn = tk.Button(
            parent,
            text=text,
            bg=bg,
            fg="white",
            activebackground=activebackground,
            activeforeground="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=padx,
            pady=4,
            command=command,
        )
        return btn

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=(6, 5, 6, 2))
        top.pack(fill="x")

        ttk.Label(top, text="COM:", style="Header.TLabel").pack(side="left")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(top, textvariable=self.port_var, width=10, state="readonly")
        self.port_combo.pack(side="left", padx=(3, 5))

        ttk.Button(top, text="Refresh", command=self.refresh_ports).pack(side="left", padx=3)

        ttk.Label(top, text="Baud:", style="Header.TLabel").pack(side="left", padx=(8, 2))
        self.baud_var = tk.StringVar(value="57600")
        ttk.Entry(top, textvariable=self.baud_var, width=8).pack(side="left", padx=(0, 5))

        self.connect_btn = ttk.Button(top, text="Connect", command=self.toggle_connection)
        self.connect_btn.pack(side="left", padx=5)

        self.start_btn = self._small_button(top, "START SAT", self.start_satellite, "#1B5E20", "#2E7D32", padx=12)
        self.start_btn.pack(side="left", padx=4)

        self.stop_btn = self._small_button(top, "STOP SAT", self.stop_satellite, "#8B5E00", "#A66F00", padx=10)
        self.stop_btn.pack(side="left", padx=4)

        self._sheets_btn = self._small_button(top, "SHEETS", self.toggle_sheets_upload, "#1A237E", "#283593", padx=10)
        self._sheets_btn.pack(side="left", padx=4)

        self._small_button(top, "MAP", self._open_map_window, "#1B4332", "#2D6A4F", padx=10).pack(side="left", padx=4)
        self._small_button(top, "EARTH", self._open_google_earth, "#3A0CA3", "#5F37D6", padx=10).pack(side="left", padx=4)
        self._test_btn = self._small_button(top, "TEST SIM", self.toggle_test_mode, "#6A1B9A", "#8E24AA", padx=10)
        self._test_btn.pack(side="left", padx=4)

        ttk.Label(top, text="F11 fullscreen / Esc exit", style="Header.TLabel").pack(side="right", padx=(10, 2))

        status = ttk.Frame(self.root, padding=(6, 0, 6, 4))
        status.pack(fill="x")
        self.status_var = tk.StringVar(value="Disconnected")
        self.packet_var = tk.StringVar(value="Packets: 0")
        self.last_rx_var = tk.StringVar(value="Last RX: -")
        self.link_var = tk.StringVar(value="Link: idle")
        self.sat_status_var = tk.StringVar(value="SAT: -")
        for var, style_name in (
            (self.status_var, "Header.TLabel"),
            (self.packet_var, "TLabel"),
            (self.last_rx_var, "TLabel"),
            (self.link_var, "Header.TLabel"),
            (self.sat_status_var, "Header.TLabel"),
        ):
            ttk.Label(status, textvariable=var, style=style_name).pack(side="left", padx=(0, 16))

        main = tk.PanedWindow(
            self.root,
            orient="horizontal",
            sashwidth=6,
            sashrelief="raised",
            bg="#0F1115",
            bd=0,
        )
        main.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        left = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, minsize=520, width=720, stretch="always")
        main.add(right, minsize=520, width=860, stretch="always")

        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        cards = ttk.Frame(left)
        cards.grid(row=0, column=0, sticky="ew")

        self.card_vars = {}
        card_items = [
            ("seq", "SEQ"),
            ("temp", "TEMP C"),
            ("pressure", "PRESS hPa"),
            ("alt", "ALT m"),
            ("peak_alt", "PEAK ALT"),
            ("deploy", "DEPLOY"),
            ("ax", "ACC X"),
            ("ay", "ACC Y"),
            ("az", "ACC Z"),
            ("gx", "GYRO X"),
            ("gy", "GYRO Y"),
            ("gz", "GYRO Z"),
            ("gps_lat", "GPS LAT"),
            ("gps_lon", "GPS LON"),
            ("gps_alt", "GNSS ALT"),
            ("gps_sats", "GPS SATS"),
        ]
        columns = 4
        for i, (key, label) in enumerate(card_items):
            frame = ttk.LabelFrame(cards, text=label, padding=(5, 4))
            frame.grid(row=i // columns, column=i % columns, padx=3, pady=3, sticky="nsew")
            var = tk.StringVar(value="-")
            ttk.Label(frame, textvariable=var, style="Value.TLabel", width=10).pack(fill="x")
            self.card_vars[key] = var
        for i in range(columns):
            cards.columnconfigure(i, weight=1)

        ts_frame = ttk.LabelFrame(left, text="Timestamp", padding=(6, 4))
        ts_frame.grid(row=1, column=0, sticky="ew", pady=(4, 4))
        self.timestamp_var = tk.StringVar(value="-")
        ttk.Label(ts_frame, textvariable=self.timestamp_var, style="Value.TLabel").pack(anchor="w", fill="x")

        left_pane = tk.PanedWindow(left, orient="vertical", sashwidth=5, sashrelief="raised", bg="#0F1115", bd=0)
        left_pane.grid(row=2, column=0, sticky="nsew")

        lower_visuals = tk.PanedWindow(left_pane, orient="horizontal", sashwidth=5, sashrelief="raised", bg="#0F1115", bd=0)
        att_frame = ttk.LabelFrame(lower_visuals, text="Attitude Indicator", padding=6)
        preview_frame = ttk.LabelFrame(lower_visuals, text="Latest Preview", padding=6)
        lower_visuals.add(att_frame, minsize=310, width=430, stretch="always")
        lower_visuals.add(preview_frame, minsize=180, width=260, stretch="always")
        left_pane.add(lower_visuals, minsize=270, height=390, stretch="always")

        self.attitude = AttitudeIndicator(att_frame, self, width=360, height=340)
        self.attitude.pack(expand=True, fill="both")
        self.attitude.draw_indicator(0.0, 0.0)

        self.preview_label = ttk.Label(preview_frame, text="No preview yet", anchor="center")
        self.preview_label.pack(fill="both", expand=True)

        console_frame = ttk.LabelFrame(left_pane, text="Console", padding=6)
        left_pane.add(console_frame, minsize=130, height=170, stretch="never")
        self.console = tk.Text(
            console_frame,
            height=7,
            wrap="word",
            bg="#11151C",
            fg="#E6EAF2",
            insertbackground="white",
            relief="flat",
            font=("Consolas", 9),
        )
        self.console.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(console_frame, orient="vertical", command=self.console.yview)
        yscroll.pack(side="right", fill="y")
        self.console.configure(yscrollcommand=yscroll.set)

        tabs = ttk.Notebook(right)
        tabs.pack(fill="both", expand=True)
        self.dashboard_tabs = tabs

        map_tab = ttk.Frame(tabs)
        orientation_tab = ttk.Frame(tabs)
        charts_tab = ttk.Frame(tabs)

        tabs.add(map_tab, text="Satellite Map")
        tabs.add(orientation_tab, text="3D Orientation")
        tabs.add(charts_tab, text="Telemetry Charts")

        map_frame = ttk.LabelFrame(map_tab, text="Satellite Map", padding=4)
        map_frame.pack(fill="both", expand=True, padx=4, pady=4)
        self._live_map = EmbeddedLiveMap(self.root, self, map_frame)
        self._live_map.set_ground_station(self._gs_lat_default, self._gs_lon_default)

        orientation_frame = ttk.LabelFrame(orientation_tab, text="3D Satellite Orientation", padding=4)
        orientation_frame.pack(fill="both", expand=True, padx=4, pady=4)
        self.orientation_3d_view = Embedded3DOrientation(self, orientation_frame)

        chart_frame = ttk.LabelFrame(charts_tab, text="Telemetry Charts", padding=4)
        chart_frame.pack(fill="both", expand=True, padx=4, pady=4)
        self.viz_window = EmbeddedTelemetryViz(self, chart_frame)

        tabs.select(map_tab)



if __name__ == "__main__":
    root = tk.Tk()
    app = FullscreenGroundStationApp(root)
    root.protocol("WM_DELETE_WINDOW", app.shutdown)
    root.mainloop()

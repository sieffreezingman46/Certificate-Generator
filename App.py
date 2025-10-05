import sys
import subprocess
import importlib
import io
import os
from pathlib import Path
import time
 

# --- Ensure required packages are installed before importing them ---

def print_start_banner():
    try:
        print("""
 ██████╗███████╗██████╗ ████████╗██╗███████╗██╗ ██████╗ █████╗ ████████╗███████╗
██╔════╝██╔════╝██╔══██╗╚══██╔══╝██║██╔════╝██║██╔════╝██╔══██╗╚══██╔══╝██╔════╝
██║     █████╗  ██████╔╝   ██║   ██║█████╗  ██║██║     ███████║   ██║   █████╗  
██║     ██╔══╝  ██╔══██╗   ██║   ██║██╔══╝  ██║██║     ██╔══██║   ██║   ██╔══╝  
╚██████╗███████╗██║  ██║   ██║   ██║██║     ██║╚██████╗██║  ██║   ██║   ███████╗
 ╚═════╝╚══════╝╚═╝  ╚═╝   ╚═╝   ╚═╝╚═╝     ╚═╝ ╚═════╝╚═╝  ╚═╝   ╚═╝   ╚══════╝
                                                                                
 ██████╗ ███████╗███╗   ██╗███████╗██████╗  █████╗ ████████╗ ██████╗ ██████╗    
██╔════╝ ██╔════╝████╗  ██║██╔════╝██╔══██╗██╔══██╗╚══██╔══╝██╔═══██╗██╔══██╗   
██║  ███╗█████╗  ██╔██╗ ██║█████╗  ██████╔╝███████║   ██║   ██║   ██║██████╔╝   
██║   ██║██╔══╝  ██║╚██╗██║██╔══╝  ██╔══██╗██╔══██║   ██║   ██║   ██║██╔══██╗   
╚██████╔╝███████╗██║ ╚████║███████╗██║  ██║██║  ██║   ██║   ╚██████╔╝██║  ██║   
 ╚═════╝ ╚══════╝╚═╝  ╚═══╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝   

Automated Certificate Generator v2.0
Smart PDF name placement with customizable UI

""")
    except Exception:
        pass


def _print_exit_and_wait():
    try:
        print("Exiting....")
        try:
            import sys as _sys
            _sys.stdout.flush()
        except Exception:
            pass
        time.sleep(3)
    except Exception:
        pass

def pre_start_update_check():
    """Print-only update check that runs before package checks.
    Uses git to compare local and remote; does not modify files or restart.
    """
    try:
        print("\n========== update: pre-flight ==========")
        print("[update] initializing...")
        repo_dir = Path(__file__).resolve().parent
        if not (repo_dir / ".git").exists():
            print("[update] repo: not a git repository -> skip")
            print("========== update: done =========\n")
            return

        def run_git(*args):
            return subprocess.run(["git", *args], cwd=str(repo_dir), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        # Fetch quietly
        print("[update] fetching: remote refs ...")
        run_git("fetch", "--all", "--prune", "--quiet")
        print("[update] fetching: ok")

        # Current branch
        res_branch = run_git("rev-parse", "--abbrev-ref", "HEAD")
        branch = (res_branch.stdout or "").strip()
        if res_branch.returncode != 0 or branch == "HEAD":
            print("[update] branch: detached or unknown -> skip")
            print("========== update: done =========\n")
            return

        res_upstream = run_git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}")
        if res_upstream.returncode == 0:
            upstream_ref = res_upstream.stdout.strip()
        else:
            upstream_ref = f"origin/{branch}"
        print(f"[update] upstream: {upstream_ref}")

        res_local = run_git("rev-parse", "HEAD")
        res_remote = run_git("rev-parse", upstream_ref)
        if res_local.returncode != 0 or res_remote.returncode != 0:
            print("[update] compare: failed -> skip")
            print("========== update: done =========\n")
            return

        local_sha = (res_local.stdout or "").strip()
        remote_sha = (res_remote.stdout or "").strip()
        if not local_sha or not remote_sha or local_sha == remote_sha:
            print("[update] status: up-to-date ✓")
            print("========== update: done =========\n")
            return

        print(f"[update] status: behind  {local_sha[:7]} -> {remote_sha[:7]}")
        print("[update] action: will apply after UI loads")
        print("========== update: done =========\n")
    except Exception:
        print("[update] error: pre-flight check skipped")
        try:
            print("========== update: done =========\n")
        except Exception:
            pass
def _get_package_version(module):
    return getattr(module, "__version__", getattr(module, "Version", "?"))

def ensure_package(package_import_name, pip_name=None):
    package_to_install = pip_name or package_import_name
    print(f"Checking {package_to_install} ...", end=" ")
    try:
        module = importlib.import_module(package_import_name)
        print(f"OK (v{_get_package_version(module)})")
        return
    except ImportError:
        print("missing")

    print(f"Installing {package_to_install} ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_to_install])
        module = importlib.import_module(package_import_name)
        print(f"Installed {package_to_install} (v{_get_package_version(module)})")
    except Exception as e:
        print(f"Failed to install {package_to_install}: {e}")
        sys.exit(1)

# Map import names to pip package names where they differ
REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("PyPDF2", "PyPDF2"),
    ("reportlab", "reportlab"),
    ("openpyxl", "openpyxl"),
    ("PIL", "Pillow"),
    ("fitz", "pymupdf"),
    ("PySide6", "PySide6"),
]

print_start_banner()
pre_start_update_check()
print("Checking required packages...")
for import_name, pip_name in REQUIRED_PACKAGES:
    ensure_package(import_name, pip_name)
print("All required packages are present.\n")

# Safe to import third-party libraries now
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
import fitz  # PyMuPDF

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QLineEdit, QPushButton, QProgressBar, 
    QFileDialog, QMessageBox, QGroupBox, QSpinBox, QFrame,
    QDialog, QSizePolicy, QScrollArea, QSplitter, QSlider
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize, QRect, QPoint
from PySide6.QtGui import (
    QPixmap, QPainter, QPen, QFont, QIcon, QFontDatabase, QBrush, QColor,
    QPolygon
)


def check_for_updates_on_startup(parent_window=None):
    """Fetch and apply updates from the git remote if the local repo is behind.
    Runs silently; if an update happens, the app will restart.
    """
    try:
        repo_dir = Path(__file__).resolve().parent
        if not (repo_dir / ".git").exists():
            return  # Not a git repo; skip

        def run_git(*args, raise_on_error=True):
            result = subprocess.run([
                "git", *args
            ], cwd=str(repo_dir), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if raise_on_error and result.returncode != 0:
                raise RuntimeError(result.stderr.strip() or "git command failed")
            return result

        # Fetch remote refs quietly
        run_git("fetch", "--all", "--prune", "--quiet", raise_on_error=False)

        # Determine current branch and upstream
        res_branch = run_git("rev-parse", "--abbrev-ref", "HEAD", raise_on_error=False)
        branch = (res_branch.stdout or "").strip()
        if res_branch.returncode != 0 or branch == "HEAD":
            return  # Detached or unknown; skip

        res_upstream = run_git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}", raise_on_error=False)
        if res_upstream.returncode == 0:
            upstream_ref = res_upstream.stdout.strip()
        else:
            upstream_ref = f"origin/{branch}"

        # Compare local and upstream commits
        res_local = run_git("rev-parse", "HEAD", raise_on_error=False)
        res_remote = run_git("rev-parse", upstream_ref, raise_on_error=False)
        if res_local.returncode != 0 or res_remote.returncode != 0:
            return

        local_sha = res_local.stdout.strip()
        remote_sha = res_remote.stdout.strip()
        if not local_sha or not remote_sha or local_sha == remote_sha:
            return  # Up to date

        # Attempt to pull updates with autostash
        pull = run_git("pull", "--rebase", "--autostash", raise_on_error=False)
        if pull.returncode != 0:
            return  # Do not disrupt the user if pull fails

        # Restart application to load updated code
        try:
            if parent_window is not None:
                QMessageBox.information(parent_window, "Updating", "An update was applied. Restarting...")
        except Exception:
            pass

        script_path = Path(__file__).resolve()
        python = sys.executable
        subprocess.Popen([python, str(script_path)] + sys.argv[1:])
        QTimer.singleShot(50, lambda: QApplication.instance().quit())
    except Exception:
        # Fail silently; never block app start due to updater
        pass

def capitalize_each_word_preserving_rest(name):
    if not name:
        return name
    words = name.split()
    capitalized_words = []
    for word in words:
        if word and word[0].islower():
            capitalized_words.append(word[0].upper() + word[1:])
        else:
            capitalized_words.append(word)
    return " ".join(capitalized_words)


class ArrowButton(QPushButton):
    def __init__(self, arrow_direction="up", parent=None):
        super().__init__(parent)
        self.arrow_direction = arrow_direction
        self.setFixedSize(10, 10) # Make it a square

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor("white"))

        w = self.width()
        h = self.height()
        
        # Centered triangle points for 10x10
        if self.arrow_direction == "up":
            points = [QPoint(w/2, h/2 - 3), QPoint(w/2 - 3, h/2 + 2), QPoint(w/2 + 3, h/2 + 2)]
        else: # down
            points = [QPoint(w/2, h/2 + 3), QPoint(w/2 - 3, h/2 - 2), QPoint(w/2 + 3, h/2 - 2)]
            
        painter.drawPolygon(QPolygon(points))


class CustomSpinBox(QWidget):
    valueChanged = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._value = 0
        self._min = 0
        self._max = 100
        self._suffix = ""

        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        self.line_edit = QLineEdit(str(self._value))
        self.line_edit.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.line_edit.textChanged.connect(self._on_text_changed)
        
        self.up_button = ArrowButton("up")
        self.down_button = ArrowButton("down")
        self.up_button.clicked.connect(self._increment)
        self.down_button.clicked.connect(self._decrement)

        button_layout = QVBoxLayout()
        button_layout.setSpacing(0)
        button_layout.setContentsMargins(0,0,0,0)
        button_layout.addWidget(self.up_button)
        button_layout.addWidget(self.down_button)

        self.layout.addWidget(self.line_edit)
        self.layout.addLayout(button_layout)

        # Make heights consistent with 2x 10px buttons
        total_h = self.up_button.height() + self.down_button.height()
        self.setFixedHeight(total_h)
        self.line_edit.setFixedHeight(total_h)

        # Make it compact (box-like) by constraining width
        self.line_edit.setFixedWidth(40)
        self.setFixedWidth(self.line_edit.width() + self.up_button.width() + 4)
        
        self.setStyleSheet("""
            QLineEdit {
                border: 1px solid #404040;
                border-right: 0px;
                border-top-left-radius: 6px;
                border-bottom-left-radius: 6px;
                padding: 0px 6px;
                background-color: #3d3d3d;
                color: #ffffff;
                font-size: 11px;
            }
            ArrowButton {
                background-color: #555555;
                border-left: 1px solid #404040;
                border-top: 0px;
                border-bottom: 0px;
                border-right: 1px solid #404040;
            }
            ArrowButton#up_button {
                border-top-right-radius: 6px;
            }
            ArrowButton#down_button {
                border-bottom-right-radius: 6px;
            }
            ArrowButton:hover {
                background-color: #666666;
            }
            ArrowButton:pressed {
                background-color: #444444;
            }
        """)
        self.up_button.setObjectName("up_button")
        self.down_button.setObjectName("down_button")

    def _increment(self):
        new_val = min(self._value + 1, self._max)
        self.setValue(new_val)

    def _decrement(self):
        new_val = max(self._value - 1, self._min)
        self.setValue(new_val)

    def _on_text_changed(self, text):
        try:
            val_str = text
            if self._suffix and val_str.endswith(self._suffix):
                val_str = val_str[: -len(self._suffix)]
            val = int(val_str) if val_str.strip() != "" else self._min
            if self._min <= val <= self._max:
                if val != self._value:
                    self._value = val
                    self.valueChanged.emit(self._value)
        except ValueError:
            # Revert to current value to avoid invalid display
            self.line_edit.blockSignals(True)
            self.line_edit.setText(f"{self._value}{self._suffix}")
            self.line_edit.blockSignals(False)

    def value(self):
        return self._value

    def setValue(self, val):
        clamped = max(self._min, min(self._max, int(val)))
        if clamped != self._value:
            self._value = clamped
            self.valueChanged.emit(self._value)
        # Always update display text
        self.line_edit.blockSignals(True)
        self.line_edit.setText(f"{self._value}{self._suffix}")
        self.line_edit.blockSignals(False)

    def setRange(self, min_val, max_val):
        self._min = int(min_val)
        self._max = int(max_val)
        self.setValue(self._value)

    def setSuffix(self, suffix):
        self._suffix = suffix or ""
        self.line_edit.blockSignals(True)
        self.line_edit.setText(f"{self._value}{self._suffix}")
        self.line_edit.blockSignals(False)


class CustomZoomBar(QWidget):
    zoom_changed = Signal(float)
    
    def __init__(self, initial_zoom=1.0, parent=None):
        super().__init__(parent)
        self.zoom_value = max(min(initial_zoom, 4.0), 0.4)  # Clamp between 0.4 and 4.0
        self.min_zoom = 0.4
        self.max_zoom = 4.0
        self.dragging = False
        self.setFixedSize(18, 100)
        self.setMouseTracking(True)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw track (thin vertical line)
        track_x = self.width() // 2
        track_top = 10
        track_bottom = self.height() - 10
        track_height = track_bottom - track_top
        
        painter.setPen(QPen(QColor(100, 100, 100), 2))
        painter.drawLine(track_x, track_top, track_x, track_bottom)
        
        # Calculate handle position
        zoom_ratio = (self.zoom_value - self.min_zoom) / (self.max_zoom - self.min_zoom)
        handle_y = track_bottom - (zoom_ratio * track_height)
        
        # Draw handle (circle)
        handle_radius = 8
        handle_rect = QRect(track_x - handle_radius, int(handle_y) - handle_radius, 
                           handle_radius * 2, handle_radius * 2)
        
        # Handle gradient
        gradient_brush = QBrush(QColor(175, 122, 197))  # Purple color matching UI
        painter.setBrush(gradient_brush)
        painter.setPen(QPen(QColor(255, 255, 255), 2))
        painter.drawEllipse(handle_rect)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.update_zoom_from_mouse(event.position().y())
            
    def mouseMoveEvent(self, event):
        if self.dragging:
            self.update_zoom_from_mouse(event.position().y())
            
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = False
            
    def update_zoom_from_mouse(self, mouse_y):
        track_top = 10
        track_bottom = self.height() - 10
        track_height = track_bottom - track_top
        
        # Clamp mouse position to track bounds
        clamped_y = max(track_top, min(track_bottom, mouse_y))
        
        # Convert to zoom ratio (inverted because top = max zoom)
        zoom_ratio = (track_bottom - clamped_y) / track_height
        
        # Convert to zoom value
        new_zoom = self.min_zoom + (zoom_ratio * (self.max_zoom - self.min_zoom))
        self.zoom_value = max(self.min_zoom, min(self.max_zoom, new_zoom))
        
        self.update()
        self.zoom_changed.emit(self.zoom_value)
        
    def set_zoom(self, zoom):
        self.zoom_value = max(self.min_zoom, min(self.max_zoom, zoom))
        self.update()


class CertificateGeneratorThread(QThread):
    progress_updated = Signal(int)
    status_updated = Signal(str)
    finished = Signal(str)
    error_occurred = Signal(str)

    def __init__(self, names_file_path, template_pdf_path, output_folder_path, 
                 font_path, font_size, y_position, name_column):
        super().__init__()
        self.names_file_path = names_file_path
        self.template_pdf_path = template_pdf_path
        self.output_folder_path = output_folder_path
        self.font_path = font_path
        self.font_size = font_size
        self.y_position = y_position
        self.name_column = name_column

    def run(self):
        try:
            if not self.names_file_path or not os.path.exists(self.names_file_path):
                raise FileNotFoundError("Names file not found")
            if not self.template_pdf_path or not os.path.exists(self.template_pdf_path):
                raise FileNotFoundError("Template PDF not found")
            if not self.output_folder_path or not os.path.isdir(self.output_folder_path):
                raise FileNotFoundError("Output folder not found")

            # Load spreadsheet (single-line to avoid any indentation ambiguity)
            names_df = (
                pd.read_excel(self.names_file_path)
                if self.names_file_path.endswith(".xlsx")
                else pd.read_csv(self.names_file_path)
            )

            if self.name_column not in names_df.columns:
                raise KeyError(f"Column '{self.name_column}' not found in sheet")

            # Register font
            if not self.font_path or not os.path.exists(self.font_path):
                raise FileNotFoundError("Font .ttf file not found")
            pdfmetrics.registerFont(TTFont("MyFont", self.font_path))

            total = len(names_df.index)
            completed = 0

            for _, row in names_df.iterrows():
                raw_name = str(row[self.name_column]).strip()
                name_value = capitalize_each_word_preserving_rest(raw_name)

                # Fresh template per certificate to avoid cumulative merges
                template_reader = PdfReader(open(self.template_pdf_path, "rb"))
                template_page = template_reader.pages[0]
                page_width = float(template_page.mediabox.width)
                page_height = float(template_page.mediabox.height)

                # Create overlay
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))
                can.setFont("MyFont", int(self.font_size))
                can.drawCentredString(page_width / 2, int(self.y_position), name_value)
                can.save()
                packet.seek(0)

                # Merge
                overlay_pdf = PdfReader(packet)
                writer = PdfWriter()
                template_page.merge_page(overlay_pdf.pages[0])
                writer.add_page(template_page)

                # Save
                output_path = os.path.join(self.output_folder_path, f"{name_value}.pdf")
                with open(output_path, "wb") as out_file:
                    writer.write(out_file)

                completed += 1
                if total > 0:
                    progress = int(completed * 100 / total)
                    self.progress_updated.emit(progress)
                self.status_updated.emit(f"Generated {completed}/{total}: {name_value}")

            self.finished.emit(f"✅ Done. Certificates saved to: {self.output_folder_path}")
        except Exception as exc:
            self.error_occurred.emit(str(exc))


class PDFPreviewDialog(QDialog):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.y_position = 300
        self.preview_scale = 1.0
        self.render_zoom = 2.0
        self.rendered_height = None
        
        self.setWindowTitle("Select Y Position")
        self.setModal(True)
        self.resize(800, 600)
        
        self.setup_ui()
        self.load_pdf()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Info label
        info_label = QLabel("Move the mouse to position the guide. Click to set Y position.")
        info_label.setStyleSheet("padding: 8px; font-weight: bold; color: #ffffff; background-color: #2d2d2d;")
        layout.addWidget(info_label)
        
        # Scroll area for PDF preview
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumSize(750, 500)
        self.scroll_area.setStyleSheet("background-color: #2d2d2d; border: 2px solid #404040;")
        
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("background-color: #343a40; border: 1px solid #495057;")
        self.preview_label.setMouseTracking(True)
        self.preview_label.mouseMoveEvent = self.on_mouse_move
        self.preview_label.mousePressEvent = self.on_mouse_click
        
        self.scroll_area.setWidget(self.preview_label)
        layout.addWidget(self.scroll_area)

    def load_pdf(self):
        try:
            doc = fitz.open(self.pdf_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(self.render_zoom, self.render_zoom))
            
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            
            # Convert PIL to QPixmap
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            pixmap = QPixmap()
            pixmap.loadFromData(img_bytes.getvalue())
            
            # Scale to fit width
            target_width = 750
            if pixmap.width() > target_width:
                pixmap = pixmap.scaledToWidth(target_width, Qt.SmoothTransformation)
            
            self.preview_scale = pixmap.width() / (pix.width)
            self.rendered_height = pix.height
            self.original_pixmap = pixmap
            
            self.preview_label.setPixmap(pixmap)
            self.preview_label.resize(pixmap.size())
            
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", str(e))

    def on_mouse_move(self, event):
        if not self.original_pixmap:
            return
            
        # Create a copy of the original pixmap
        pixmap = self.original_pixmap.copy()
        painter = QPainter(pixmap)
        
        # Draw guide line
        pen = QPen(Qt.black, 2)
        painter.setPen(pen)
        y = event.position().y()
        painter.drawLine(0, y, pixmap.width(), y)
        
        # Calculate and draw Y position text
        if self.preview_scale and self.rendered_height:
            canvas_y = max(min(y, int(self.rendered_height * self.preview_scale)), 0)
            rendered_y = canvas_y / self.preview_scale
            pdf_y_points = (self.rendered_height - rendered_y) / self.render_zoom
            
            # Draw Y position text
            font = QFont("Segoe UI", 12, QFont.Bold)
            painter.setFont(font)
            painter.setPen(QPen(Qt.black, 1))
            text = f"Y: {int(pdf_y_points)}"
            text_y = max(y - 15, 15)
            text_x = pixmap.width() // 2
            painter.drawText(text_x - 25, text_y, text)
        
        painter.end()
        self.preview_label.setPixmap(pixmap)

    def on_mouse_click(self, event):
        if self.preview_scale and self.rendered_height:
            canvas_y = max(min(event.position().y(), int(self.rendered_height * self.preview_scale)), 0)
            rendered_y = canvas_y / self.preview_scale
            pdf_y_points = (self.rendered_height - rendered_y) / self.render_zoom
            self.y_position = int(pdf_y_points)
            self.accept()


class CertificateGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.names_file_path = ""
        self.template_pdf_path = ""
        self.output_folder_path = ""
        self.font_path = ""
        self.font_size = 30
        self.y_position = 300
        self.name_column = "Name"
        self.preview_zoom = 1.0
        self._panning = False
        self._pan_start = None
        self._fit_on_next_load = False
        
        self.setup_ui()
        self.apply_styles()
        self.setup_defaults()

    def setup_ui(self):
        self.setWindowTitle("Certificate Generator")
        self.setMinimumSize(1000, 700)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(16)
        main_layout.setContentsMargins(16, 16, 16, 16)
        
        # Left panel - Inputs and Options
        left_panel = QWidget()
        left_panel.setMaximumWidth(400)
        left_layout = QVBoxLayout(left_panel)
        
        # File Inputs Group
        inputs_group = QGroupBox("File Inputs")
        inputs_layout = QGridLayout(inputs_group)
        
        # Names file
        inputs_layout.addWidget(QLabel("Names file:"), 0, 0)
        self.names_edit = QLineEdit()
        self.names_edit.setPlaceholderText("Select Excel or CSV file...")
        inputs_layout.addWidget(self.names_edit, 0, 1)
        self.names_btn = QPushButton("Browse")
        self.names_btn.clicked.connect(self.browse_names)
        inputs_layout.addWidget(self.names_btn, 0, 2)
        
        # Template PDF
        inputs_layout.addWidget(QLabel("Template PDF:"), 1, 0)
        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("Select PDF template...")
        inputs_layout.addWidget(self.template_edit, 1, 1)
        self.template_btn = QPushButton("Browse")
        self.template_btn.clicked.connect(self.browse_template)
        inputs_layout.addWidget(self.template_btn, 1, 2)
        
        # Output folder
        inputs_layout.addWidget(QLabel("Output folder:"), 2, 0)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Select output folder...")
        inputs_layout.addWidget(self.output_edit, 2, 1)
        self.output_btn = QPushButton("Browse")
        self.output_btn.clicked.connect(self.browse_output)
        inputs_layout.addWidget(self.output_btn, 2, 2)
        
        # Font file
        inputs_layout.addWidget(QLabel("Font file:"), 3, 0)
        self.font_edit = QLineEdit()
        self.font_edit.setPlaceholderText("Select TTF font file...")
        inputs_layout.addWidget(self.font_edit, 3, 1)
        self.font_btn = QPushButton("Browse")
        self.font_btn.clicked.connect(self.browse_font)
        inputs_layout.addWidget(self.font_btn, 3, 2)
        
        inputs_layout.setColumnStretch(1, 1)
        left_layout.addWidget(inputs_group)
        
        # Options Group
        options_group = QGroupBox("Options")
        options_layout = QGridLayout(options_group)
        
        # Font size
        options_layout.addWidget(QLabel("Font size:"), 0, 0)
        self.font_size_edit = QLineEdit("30")
        self.font_size_edit.setPlaceholderText("e.g. 30")
        self.font_size_edit.textChanged.connect(self.on_font_size_text_changed)
        options_layout.addWidget(self.font_size_edit, 0, 1)
        
        # Y position
        options_layout.addWidget(QLabel("Y position:"), 1, 0)
        self.y_pos_edit = QLineEdit("300")
        self.y_pos_edit.setPlaceholderText("e.g. 300")
        self.y_pos_edit.textChanged.connect(self.on_y_position_text_changed)
        options_layout.addWidget(self.y_pos_edit, 1, 1)
        
        self.reconfigure_btn = QPushButton("Reconfigure Y")
        self.reconfigure_btn.clicked.connect(self.reconfigure_y)
        options_layout.addWidget(self.reconfigure_btn, 1, 2)
        
        # Name column
        options_layout.addWidget(QLabel("Name column:"), 2, 0)
        self.name_col_edit = QLineEdit("Name")
        self.name_col_edit.textChanged.connect(lambda t: setattr(self, 'name_column', t))
        options_layout.addWidget(self.name_col_edit, 2, 1, 1, 2)

        # Removed Preview font tweak (%)
        # options_layout.addWidget(QLabel("Preview font %:"), 3, 0)
        # self.preview_tweak_spin = CustomSpinBox()
        # self.preview_tweak_spin.setRange(50, 150)
        # self.preview_tweak_spin.setValue(self.preview_tweak_percent)
        # self.preview_tweak_spin.setSuffix("%")
        # self.preview_tweak_spin.valueChanged.connect(self.on_preview_tweak_changed)
        # options_layout.addWidget(self.preview_tweak_spin, 3, 1)
        
        left_layout.addWidget(options_group)
        
        # Actions Group
        actions_group = QGroupBox("Actions")
        actions_layout = QVBoxLayout(actions_group)
        
        # Generate button
        self.generate_btn = QPushButton("Generate Certificates")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.clicked.connect(self.generate_certificates)
        actions_layout.addWidget(self.generate_btn)
        
        # Open output folder button
        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        actions_layout.addWidget(self.open_folder_btn)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        actions_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("Ready")
        self.status_label.setWordWrap(True)
        actions_layout.addWidget(self.status_label)
        
        left_layout.addWidget(actions_group)
        left_layout.addStretch()
        
        # Right panel - Preview
        right_panel = QGroupBox("Preview")
        right_layout = QVBoxLayout(right_panel)
        # Bottom-right zoom controls: vertical slider and magnifier buttons
        
        # Scroll area for preview
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setMinimumHeight(400)
        
        self.preview_widget = QLabel("PDF preview will appear here when template is selected")
        self.preview_widget.setAlignment(Qt.AlignCenter)
        self.preview_widget.setStyleSheet("color: #adb5bd; font-style: italic; background-color: #343a40; border: 2px dashed #495057;")
        self.preview_widget.setMouseTracking(True)
        # Enable panning by dragging
        self.preview_widget.mousePressEvent = self.on_preview_press
        self.preview_widget.mouseMoveEvent = self.on_preview_move
        self.preview_widget.mouseReleaseEvent = self.on_preview_release
        # Enable pinch-to-zoom with wheel events
        self.preview_widget.wheelEvent = self.on_preview_wheel
        
        self.preview_scroll.setWidget(self.preview_widget)
        
        # Store zoom overlay reference for positioning
        self.zoom_overlay = None
        
        # Create a container for preview with overlay controls
        preview_container = QWidget()
        preview_container_layout = QVBoxLayout(preview_container)
        preview_container_layout.setContentsMargins(0, 0, 0, 0)
        
        # Add the scroll area to the container
        preview_container_layout.addWidget(self.preview_scroll, 1)
        
        # Create zoom controls overlay (positioned inside preview panel)
        self.zoom_overlay = QWidget(self.preview_scroll.viewport())
        self.zoom_overlay.setFixedSize(70, 110)
        self.zoom_overlay.move(10, -1)  # Bottom-right corner inside preview (will be repositioned in resizeEvent)
        # Keep the overlay non-painted without forcing children to be transparent
        self.zoom_overlay.setAutoFillBackground(False)
        self.zoom_overlay.setObjectName("zoom_overlay")
        # Only the overlay is transparent; children (buttons/slider) keep their own backgrounds
        self.zoom_overlay.setStyleSheet("#zoom_overlay { background-color: transparent; border: 0px; }")

        zoom_layout = QHBoxLayout(self.zoom_overlay)
        zoom_layout.setContentsMargins(3, 3, 3, 3)
        zoom_layout.setSpacing(8)
        
        # Custom vertical drag bar
        self.zoom_drag_bar = CustomZoomBar(self.preview_zoom)
        self.zoom_drag_bar.zoom_changed.connect(self.on_zoom_changed)
        zoom_layout.addWidget(self.zoom_drag_bar)
        
        # Buttons column (on the right) - tighter spacing
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(0)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        
        # Zoom in button - square with plus
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setFixedSize(25, 25)
        self.zoom_in_btn.clicked.connect(self.on_zoom_in)
        btn_layout.addWidget(self.zoom_in_btn)
        
        # Zoom out button - square with minus
        self.zoom_out_btn = QPushButton("-")
        self.zoom_out_btn.setFixedSize(25, 25)
        self.zoom_out_btn.clicked.connect(self.on_zoom_out)
        btn_layout.addWidget(self.zoom_out_btn)
        
        zoom_layout.addLayout(btn_layout)
        
        # Hide zoom controls until a template is chosen
        self.zoom_overlay.setVisible(False)
        
        right_layout.addWidget(preview_container, 1)
        
        # Add panels to main layout
        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel, 1)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
                color: #ffffff;
            }
            
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
            }
            
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #ffffff;
                border: 2px solid #404040;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
                background-color: #2d2d2d;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #2d2d2d;
                color: #ffffff;
            }
            
            QLabel {
                color: #e0e0e0;
                font-size: 11px;
                background-color: transparent;
            }
            
            QLineEdit {
                border: 2px solid #404040;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 11px;
                background-color: #3d3d3d;
                color: #ffffff;
            }
            
            QLineEdit:focus {
                border-color: #8e44ad;
                outline: none;
            }
            
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #9b59b6, stop:0.5 #8e44ad, stop:1 #7d3c98);
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 18px;
                font-size: 11px;
                font-weight: bold;
                min-width: 80px;
            }
            
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #af7ac5, stop:0.5 #a569bd, stop:1 #9b59b6);
            }
            
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #7d3c98, stop:0.5 #6c3483, stop:1 #5b2c6f);
            }
            
            QPushButton:disabled {
                background-color: #555555;
                color: #888888;
            }
            
            QPushButton#generate_btn {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #9b59b6, stop:0.5 #8e44ad, stop:1 #7d3c98);
                font-size: 13px;
                font-weight: bold;
                padding: 12px 20px;
            }
            
            QPushButton#generate_btn:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #af7ac5, stop:0.5 #a569bd, stop:1 #9b59b6);
            }
            
            QPushButton#open_folder_btn {
                background-color: #1a1a1a;
                color: white;
            }
            
            QPushButton#open_folder_btn:hover {
                background-color: #2a2a2a;
            }
            
            QPushButton#zoom_in_btn, QPushButton#zoom_out_btn {
                background-color: #2d2d2d;
                color: #aaaaaa;
                border: 1px solid #404040;
                border-radius: 4px;
                font-size: 16px;
                font-weight: normal;
                font-family: "Segoe UI", "Arial", sans-serif;
                padding: 0px;
                min-width: 0px;
            }
            
            QPushButton#zoom_in_btn:hover, QPushButton#zoom_out_btn:hover {
                background-color: #3d3d3d;
                color: #cccccc;
            }
            
            QPushButton#zoom_in_btn:pressed, QPushButton#zoom_out_btn:pressed {
                background-color: #1d1d1d;
                color: #888888;
            }
            
            
            QProgressBar {
                border: 2px solid #404040;
                border-radius: 6px;
                background-color: #3d3d3d;
                text-align: center;
                font-size: 11px;
                font-weight: bold;
                color: #ffffff;
            }
            
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #9b59b6, stop:1 #8e44ad);
                border-radius: 4px;
            }
            
            QScrollArea {
                border: 2px solid #404040;
                border-radius: 6px;
                background-color: #2d2d2d;
            }
            
            QScrollBar:vertical, QScrollBar:horizontal {
                background-color: #2d2d2d;
                border: none;
                border-radius: 6px;
            }
            
            QScrollBar:vertical {
                width: 12px;
                margin: 0px;
            }
            
            QScrollBar:horizontal {
                height: 12px;
                margin: 0px;
            }
            
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #9b59b6, stop:0.5 #8e44ad, stop:1 #7d3c98);
                border: none;
                border-radius: 6px;
                min-height: 20px;
                min-width: 20px;
            }
            
            QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #af7ac5, stop:0.5 #a569bd, stop:1 #9b59b6);
            }
            
            QScrollBar::handle:vertical:pressed, QScrollBar::handle:horizontal:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #7d3c98, stop:0.5 #6c3483, stop:1 #5b2c6f);
            }
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none;
                border: none;
                width: 0px;
                height: 0px;
            }
            
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """)
        
        # Set object names for specific styling
        self.generate_btn.setObjectName("generate_btn")
        self.open_folder_btn.setObjectName("open_folder_btn")
        self.zoom_in_btn.setObjectName("zoom_in_btn")
        self.zoom_out_btn.setObjectName("zoom_out_btn")

    def setup_defaults(self):
        # Set default font path if exists
        default_font_path = os.path.join(os.getcwd(), "Montserrat-Medium.ttf")
        if os.path.exists(default_font_path):
            self.font_edit.setText(default_font_path)
            self.font_path = default_font_path

    def browse_names(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Names File", "", 
            "Excel files (*.xlsx);;CSV files (*.csv);;All files (*)"
        )
        if file_path:
            self.names_edit.setText(file_path)
            self.names_file_path = file_path
            # Refresh preview if template is already loaded
            if self.template_pdf_path:
                self.load_preview()

    def browse_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Template PDF", "", 
            "PDF files (*.pdf);;All files (*)"
        )
        if file_path:
            self.template_edit.setText(file_path)
            self.template_pdf_path = file_path
            # Load preview
            self._fit_on_next_load = True
            self.load_preview()
            # Show zoom controls now that a template is selected and position them
            try:
                self.zoom_overlay.setVisible(True)
                QTimer.singleShot(0, self.position_zoom_overlay)
            except Exception:
                pass
            # Auto-open Y position selector
            self.reconfigure_y()

    def browse_output(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.output_edit.setText(folder_path)
            self.output_folder_path = folder_path

    def browse_font(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Font File", "", 
            "TrueType fonts (*.ttf);;All files (*)"
        )
        if file_path:
            self.font_edit.setText(file_path)
            self.font_path = file_path
            # Refresh preview if template is already loaded
            if self.template_pdf_path:
                self.load_preview()

    def on_font_size_changed(self, value):
        self.font_size = value
        # Refresh preview if template is already loaded
        if self.template_pdf_path:
            self.load_preview()

    def on_font_size_text_changed(self, text):
        """Live update font size as user types"""
        try:
            value = int(text) if text.strip() else 30
            if 8 <= value <= 200:  # Valid range
                self.font_size = value
                if self.template_pdf_path:
                    self.load_preview()
        except ValueError:
            pass  # Ignore invalid input while typing

    def on_y_position_changed(self, value):
        self.y_position = value
        # Refresh preview if template is already loaded
        if self.template_pdf_path:
            self.load_preview()

    def on_y_position_text_changed(self, text):
        """Live update Y position as user types"""
        try:
            value = int(text) if text.strip() else 300
            if 0 <= value <= 2000:  # Valid range
                self.y_position = value
                if self.template_pdf_path:
                    self.load_preview()
        except ValueError:
            pass  # Ignore invalid input while typing

    # Removed on_preview_tweak_changed as preview tweak control was removed
    # def on_preview_tweak_changed(self, value):
    #     self.preview_tweak_percent = int(value)
    #     if self.template_pdf_path:
    #         self.load_preview()

    def load_preview(self):
        if not self.template_pdf_path or not os.path.exists(self.template_pdf_path):
            return
        
        # Position zoom overlay at bottom-left after preview loads
        if self.zoom_overlay and self.preview_scroll:
            QTimer.singleShot(100, self.position_zoom_overlay)
            
        try:
            doc = fitz.open(self.template_pdf_path)
            page = doc.load_page(0)
            # Fit-to-width on first load after selecting a template
            if self._fit_on_next_load:
                try:
                    viewport_width = max(self.preview_scroll.viewport().width(), 1)
                except Exception:
                    viewport_width = 800
                pdf_width_pts = float(page.rect.width)
                target_zoom = viewport_width / max(pdf_width_pts, 1.0)
                self.preview_zoom = max(min(target_zoom, 4.0), 0.4)
                # sync drag bar if present
                try:
                    self.zoom_drag_bar.set_zoom(self.preview_zoom)
                except Exception:
                    pass
                self._fit_on_next_load = False

            pix = page.get_pixmap(matrix=fitz.Matrix(self.preview_zoom, self.preview_zoom))
            
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            
            # Convert PIL to QPixmap
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            pixmap = QPixmap()
            pixmap.loadFromData(img_bytes.getvalue())
            
            # Add first name preview if names file and font file are selected
            if self.names_file_path and os.path.exists(self.names_file_path) and self.font_path and os.path.exists(self.font_path):
                try:
                    if self.names_file_path.endswith(".xlsx"):
                        names_df = pd.read_excel(self.names_file_path)
                    else:
                        names_df = pd.read_csv(self.names_file_path)
                    
                    if self.name_column in names_df.columns and len(names_df) > 0:
                        first_name = str(names_df.iloc[0][self.name_column]).strip()
                        first_name = capitalize_each_word_preserving_rest(first_name)
                        # Convert to QPixmap respecting template DPI
                        # Use the actual font TTF chosen by the user
                        temp_font_id = QFontDatabase.addApplicationFont(self.font_path)
                        font_family = "Arial"
                        if temp_font_id != -1:
                            families = QFontDatabase.applicationFontFamilies(temp_font_id)
                            if families:
                                font_family = families[0]

                        painter = QPainter(pixmap)
                        # Determine scale between PDF points and preview pixels
                        pdf_width_pts = float(page.rect.width)
                        pdf_height_pts = float(page.rect.height)
                        scale_x = pixmap.width() / pdf_width_pts
                        scale_y = pixmap.height() / pdf_height_pts

                        # Map PDF point size to pixels using page scale; keep independent of zoom artifacts
                        preview_font_px = max(int(self.font_size * scale_y), 1)
                        font = QFont(font_family)
                        font.setPixelSize(preview_font_px)
                        painter.setFont(font)
                        painter.setPen(QPen(Qt.black, 2))

                        # Centered X at half width; Y from bottom to top mapping using scale
                        center_x = pixmap.width() // 2
                        preview_y = pixmap.height() - int(self.y_position * scale_y)

                        # Draw centered text
                        text_rect = painter.fontMetrics().boundingRect(first_name)
                        painter.drawText(center_x - text_rect.width() // 2, preview_y, first_name)
                        painter.end()
                except:
                    pass  # Ignore errors in preview name rendering
            
            # Do not auto-scale; let zoom/scroll area control size
            self.preview_widget.setPixmap(pixmap)
            self.preview_widget.resize(pixmap.size())
            self.preview_widget.setStyleSheet("")  # Remove placeholder styling
            
        except Exception as e:
            self.preview_widget.setText(f"Error loading preview: {str(e)}")
            self.preview_widget.setStyleSheet("color: #e74c3c; font-style: italic; background-color: #343a40;")

    def reconfigure_y(self):
        if not self.template_pdf_path:
            QMessageBox.information(self, "Preview", "Please select a template PDF first.")
            return
            
        dialog = PDFPreviewDialog(self.template_pdf_path, self)
        if dialog.exec() == QDialog.Accepted:
            self.y_position = dialog.y_position
            try:
                self.y_pos_edit.setText(str(self.y_position))
            except Exception:
                pass
            # Refresh preview with new Y position
            self.load_preview()

    def generate_certificates(self):
        # Validate inputs
        if not all([self.names_file_path, self.template_pdf_path, 
                   self.output_folder_path, self.font_path]):
            QMessageBox.warning(self, "Missing Information", 
                              "Please fill in all required fields.")
            return
        
        # Disable generate button and show progress
        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting generation...")
        
        # Start generation thread
        self.generator_thread = CertificateGeneratorThread(
            self.names_file_path, self.template_pdf_path, self.output_folder_path,
            self.font_path, self.font_size, self.y_position, self.name_column
        )
        
        self.generator_thread.progress_updated.connect(self.progress_bar.setValue)
        self.generator_thread.status_updated.connect(self.status_label.setText)
        self.generator_thread.finished.connect(self.on_generation_finished)
        self.generator_thread.error_occurred.connect(self.on_generation_error)
        
        self.generator_thread.start()

    def on_generation_finished(self, message):
        self.status_label.setText(message)
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.information(self, "Success", "Certificates generated successfully!")

    def on_generation_error(self, error_message):
        self.status_label.setText(f"❌ Error: {error_message}")
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", error_message)

    def on_zoom_in(self):
        self.preview_zoom = min(self.preview_zoom * 1.2, 4.0)
        self.zoom_drag_bar.set_zoom(self.preview_zoom)
        self.load_preview()

    def on_zoom_out(self):
        self.preview_zoom = max(self.preview_zoom / 1.2, 0.4)
        self.zoom_drag_bar.set_zoom(self.preview_zoom)
        self.load_preview()

    def on_zoom_changed(self, zoom_value):
        self.preview_zoom = zoom_value
        self.load_preview()

    def on_preview_wheel(self, event):
        """Handle mouse wheel events for pinch-to-zoom"""
        if event.modifiers() & Qt.ControlModifier:
            # Ctrl + wheel for zooming
            delta = event.angleDelta().y()
            if delta > 0:
                # Zoom in
                self.preview_zoom = min(self.preview_zoom * 1.1, 4.0)
            else:
                # Zoom out
                self.preview_zoom = max(self.preview_zoom / 1.1, 0.4)
            
            # Update zoom controls
            try:
                self.zoom_drag_bar.set_zoom(self.preview_zoom)
            except Exception:
                pass
            
            # Reload preview with new zoom
            self.load_preview()
            
            # Accept the event to prevent scrolling
            event.accept()
        else:
            # Let normal scrolling happen
            event.ignore()

    # Panning handlers
    def on_preview_press(self, event):
        if event.buttons() & Qt.LeftButton:
            self._panning = True
            self._pan_start = event.position().toPoint()
            self.preview_scroll.setCursor(Qt.ClosedHandCursor)

    def on_preview_move(self, event):
        if self._panning and self._pan_start is not None:
            current = event.position().toPoint()
            delta = current - self._pan_start
            bar_h = self.preview_scroll.horizontalScrollBar()
            bar_v = self.preview_scroll.verticalScrollBar()
            bar_h.setValue(bar_h.value() - delta.x())
            bar_v.setValue(bar_v.value() - delta.y())
            self._pan_start = current

    def on_preview_release(self, event):
        self._panning = False
        self.preview_scroll.unsetCursor()

    def position_zoom_overlay(self):
        """Position the zoom overlay at the bottom-right of the preview scroll area"""
        if self.zoom_overlay and self.preview_scroll:
            viewport = self.preview_scroll.viewport()
            x = viewport.width() - self.zoom_overlay.width() - 10  # Right margin
            y = viewport.height() - self.zoom_overlay.height() - 10  # Bottom margin
            self.zoom_overlay.move(x, y)

    def open_output_folder(self):
        if not self.output_folder_path or not os.path.isdir(self.output_folder_path):
            QMessageBox.information(self, "Open Folder", 
                                  "Please select a valid output folder first.")
            return
        
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.output_folder_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.output_folder_path])
            else:
                subprocess.Popen(["xdg-open", self.output_folder_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open folder: {str(e)}")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Modern Qt style
    
    # Set application properties
    app.setApplicationName("Certificate Generator")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Certificate Tools")
    
    window = CertificateGeneratorApp()
    window.show()
    
    # Check for updates shortly after the UI shows, so prompts have a parent
    try:
        QTimer.singleShot(200, lambda: check_for_updates_on_startup(window))
    except Exception:
        pass
    
    # Print on exit
    try:
        app.aboutToQuit.connect(_print_exit_and_wait)
    except Exception:
        pass
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
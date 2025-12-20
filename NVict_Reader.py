# -*- coding: utf-8 -*-
"""
NVict Reader (Modern UI Style) - Finale Versie
Gebaseerd op de UI-stijl van NV Sync
Ontwikkeld door NVict Service

Website: www.nvict.nl
Versie: 1.6
"""

import sys
import os
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageOps, ImageDraw
import io
import tempfile
import subprocess
import platform
from datetime import datetime
import urllib.request
import urllib.error
import json
import threading
import socket
import time

# Applicatie versie
APP_VERSION = "1.6"
UPDATE_CHECK_URL = "https://www.nvict.nl/software/updates/nvict_reader_version.json"

try:
    import winreg
except ImportError:
    winreg = None

# ====================================================================
# DEFAULT PDF HANDLER - Set as Default Functionality
# ====================================================================

class DefaultPDFHandler:
    """Handles setting NVict Reader as default PDF viewer"""
    
    @staticmethod
    def is_default_pdf_handler():
        """Check if NVict Reader is currently the default PDF handler"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, 
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice",
                0, 
                winreg.KEY_READ
            )
            prog_id, _ = winreg.QueryValueEx(key, "ProgId")
            winreg.CloseKey(key)
            return "NVictReader" in prog_id or "Applications\\NVict Reader.exe" in prog_id
        except:
            return False
    
    @staticmethod
    def open_windows_default_apps_pdf():
        """Open Windows Settings directly to PDF file association"""
        try:
            # Probeert direct naar de .pdf instelling te gaan (Windows 10/11)
            subprocess.run(['start', 'ms-settings:defaultapps'], shell=True)
            return True
        except:
            return False

    @staticmethod
    def register_open_with():
        """
        Registreer NVict Reader in het register.
        CHECK: Als Inno Setup het al in HKLM heeft gezet, doen we hier NIETS om dubbele items te voorkomen.
        """
        try:
            # 1. CHECK: Is de app al globaal geïnstalleerd via Inno Setup?
            # We kijken of de ProgID in HKEY_LOCAL_MACHINE bestaat.
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\NVictReader.PDF", 0, winreg.KEY_READ)
                winreg.CloseKey(key)
                # Gevonden! De installer heeft zijn werk gedaan.
                # Wij doen niets in Python om duplicaten te voorkomen.
                return True
            except OSError:
                # Niet gevonden in HKLM, dus we draaien waarschijnlijk portable.
                # Ga door met registreren in HKCU.
                pass

            # ---------------------------------------------------------
            # Code voor Portable Versie (Schrijft naar HKEY_CURRENT_USER)
            # ---------------------------------------------------------
            
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(sys.argv[0])
            
            prog_id = "NVictReader.PDF"
            hkcu = winreg.HKEY_CURRENT_USER
            
            # RegisteredApplications pad
            cap_path = f"Software\\NVict Service\\NVict Reader\\Capabilities"
            
            # Maak de Capabilities sleutels
            key = winreg.CreateKey(hkcu, cap_path)
            winreg.SetValueEx(key, "ApplicationName", 0, winreg.REG_SZ, "NVict Reader")
            winreg.SetValueEx(key, "ApplicationDescription", 0, winreg.REG_SZ, "NVict Reader PDF Viewer")
            winreg.CloseKey(key)
            
            # FileAssociations binnen Capabilities
            key = winreg.CreateKey(hkcu, f"{cap_path}\\FileAssociations")
            winreg.SetValueEx(key, ".pdf", 0, winreg.REG_SZ, prog_id)
            winreg.CloseKey(key)
            
            # Voeg toe aan RegisteredApplications
            key = winreg.CreateKey(hkcu, "Software\\RegisteredApplications")
            winreg.SetValueEx(key, "NVictReader", 0, winreg.REG_SZ, cap_path)
            winreg.CloseKey(key)
            
            # De ProgID
            classes_path = f"Software\\Classes\\{prog_id}"
            
            key = winreg.CreateKey(hkcu, classes_path)
            winreg.SetValue(key, "", winreg.REG_SZ, "NVict Reader PDF")
            winreg.CloseKey(key)
            
            # Gebruik PDF_File_icon.ico voor Windows Verkenner, anders exe zelf
            key = winreg.CreateKey(hkcu, f"{classes_path}\\DefaultIcon")
            icon_dir = os.path.dirname(exe_path)
            pdf_icon_path = os.path.join(icon_dir, "PDF_File_icon.ico")
            if os.path.exists(pdf_icon_path):
                winreg.SetValue(key, "", winreg.REG_SZ, f'"{pdf_icon_path}",0')
            else:
                winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}",0')
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\open\\command")
            winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}" "%1"')
            winreg.CloseKey(key)
            
            # Voeg "Afdrukken" toe aan rechtsklik menu
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print")
            winreg.SetValue(key, "", winreg.REG_SZ, "Afdrukken")
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print\\command")
            winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}" --print "%1"')
            winreg.CloseKey(key)
            
            # Koppel .pdf aan ProgID
            key = winreg.CreateKey(hkcu, "Software\\Classes\\.pdf\\OpenWithProgids")
            winreg.SetValueEx(key, prog_id, 0, winreg.REG_NONE, b'')
            winreg.CloseKey(key)

            # Refresh de shell iconen
            try:
                import ctypes
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, 0, 0)
            except:
                pass
            
            return True
        except Exception as e:
            print(f"Error registering: {e}")
            return False

    @staticmethod
    def prompt_set_as_default(parent):
        """Vraagt de gebruiker om de app als standaard in te stellen"""
        # Eerst registreren in het register om zeker te zijn dat we in de lijst staan
        DefaultPDFHandler.register_open_with()
        
        msg = (
            "Om NVict Reader als standaard in te stellen, opent Windows nu het instellingen menu.\n\n"
            "1. Zoek '.pdf' in de lijst of klik op de huidige standaard app.\n"
            "2. Selecteer 'NVict Reader' in de lijst.\n"
            "3. Klik op 'Als standaard instellen'.\n\n"
            "Wilt u de instellingen nu openen?"
        )
        
        if messagebox.askyesno("Standaard App Instellen", msg, parent=parent):
            DefaultPDFHandler.open_windows_default_apps_pdf()

    @staticmethod
    def show_first_run_dialog(parent):
        """Toon dialoog bij eerste keer opstarten"""
        # Check of we al standaard zijn
        if DefaultPDFHandler.is_default_pdf_handler():
            return "already_default"

        msg = (
            "Welkom bij NVict Reader!\n\n"
            "Wilt u NVict Reader instellen als uw standaard PDF programma?\n"
            "Dit kunt u later altijd nog wijzigen via Instellingen."
        )
        
        # Maak een custom dialoog of gebruik standaard messagebox
        # We gebruiken hier askyesnocancel voor Ja / Nee / Nooit meer vragen
        # Maar tkinter heeft geen 'Nooit', dus we gebruiken simpel Ja/Nee
        
        # Omdat de return value in jouw code "never" verwacht, simuleren we dat:
        # Je kunt hier een custom dialoog bouwen, maar voor nu volstaat askyesno.
        
        if messagebox.askyesno("Welkom", msg, parent=parent):
            DefaultPDFHandler.prompt_set_as_default(parent)
            return "yes"
        else:
            return "no"

# ====================================================================

class SingleInstance:
    """Zorgt ervoor dat er maar één instance van de applicatie draait."""
    def __init__(self, port=52847):
        self.port = port
        self.sock = None
        self.server_thread = None
        self.app = None
        self.running = False
        
    def is_already_running(self):
        """Controleer of er al een instance draait."""
        try:
            # Probeer te connecteren met bestaande instance
            test_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_sock.settimeout(1)
            test_sock.connect(('127.0.0.1', self.port))
            test_sock.close()
            return True
        except (socket.error, socket.timeout):
            return False
    
    def send_to_existing_instance(self, filepath):
        """Stuur bestandspad naar bestaande instance."""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            sock.connect(('127.0.0.1', self.port))
            sock.sendall(filepath.encode('utf-8'))
            sock.close()
            return True
        except Exception as e:
            print(f"Fout bij versturen naar bestaande instance: {e}")
            return False
    
    def start_server(self, app):
        """Start socket server om berichten van andere instances te ontvangen."""
        self.app = app
        self.running = True
        
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.sock.bind(('127.0.0.1', self.port))
            self.sock.listen(5)
            self.sock.settimeout(1)
            
            # Start server thread
            self.server_thread = threading.Thread(target=self._server_loop, daemon=True)
            self.server_thread.start()
            
            return True
        except Exception as e:
            print(f"Kon single instance server niet starten: {e}")
            return False
    
    def _server_loop(self):
        """Luister naar berichten van andere instances."""
        while self.running:
            try:
                conn, addr = self.sock.accept()
                conn.settimeout(2)
                
                # Ontvang bestandspad
                data = conn.recv(4096).decode('utf-8')
                conn.close()
                
                if data and self.app:
                    # Open bestand in bestaande instance (in main thread)
                    self.app.root.after(0, lambda path=data: self.app.add_new_tab(path))
                    # Breng window naar voren
                    self.app.root.after(10, lambda: self.app.root.lift())
                    self.app.root.after(10, lambda: self.app.root.focus_force())
                    
            except socket.timeout:
                continue
            except Exception as e:
                if self.running:
                    print(f"Fout in server loop: {e}")
    
    def stop(self):
        """Stop de server."""
        self.running = False
        if self.sock:
            try:
                self.sock.close()
            except:
                pass

def get_resource_path(relative_path):
    """Geef het absolute pad naar resource bestanden (werkt met PyInstaller)"""
    try:
        # PyInstaller maakt een temp folder en slaat path op in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_settings_path():
    """Geef pad naar settings bestand in gebruikers map"""
    if platform.system() == "Windows":
        app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
        settings_dir = os.path.join(app_data, 'NVict PDF Reader')
    else:
        settings_dir = os.path.join(os.path.expanduser('~'), '.nvict_pdf_reader')
    
    # Maak directory als die niet bestaat
    os.makedirs(settings_dir, exist_ok=True)
    return os.path.join(settings_dir, 'settings.json')

class Theme:
    """Bevat de kleurenschema's voor lichte en donkere thema's."""
    LIGHT = {
        "BG_PRIMARY": "#f3f3f3", "BG_SECONDARY": "#ffffff",
        "TEXT_PRIMARY": "#1c1c1c", "TEXT_SECONDARY": "#737373",
        "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
        "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
        "SELECTION_COLOR": "#FFD700"  # Goud/geel voor betere zichtbaarheid
    }
    DARK = {
        "BG_PRIMARY": "#1e1e1e", "BG_SECONDARY": "#2d2d2d",
        "TEXT_PRIMARY": "#f0f0f0", "TEXT_SECONDARY": "#a0a0a0",
        "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
        "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
        "SELECTION_COLOR": "#FFD700"  # Goud/geel
    }
    FONT_MAIN = ("Segoe UI Variable", 10)
    FONT_HEADING = ("Segoe UI Variable", 12, "bold")
    FONT_SMALL = ("Segoe UI Variable", 9)

class PDFTab(tk.Frame):
    """Een enkel tabblad dat een PDF-document beheert."""
    def __init__(self, master, file_path, theme, password=None):
        super().__init__(master, bg=theme["BG_PRIMARY"])
        self.theme = theme
        
        # Document state
        self.file_path = file_path
        self.pdf_document = fitz.open(file_path)
        
        # Authenticeer met wachtwoord indien nodig
        if password and self.pdf_document.needs_pass:
            self.pdf_document.authenticate(password)
        
        self.current_page = 0
        self.zoom_level = 1.0
        self.zoom_mode = "fit_width"
        
        # UI elements
        self.canvas = tk.Canvas(self, bg=theme["BG_PRIMARY"], relief="flat", bd=0, 
                               highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Selection data
        self.text_words = []
        self.drag_start = None
        self.drag_rect = None
        self.selection_rects = []
        self.selected_text = ""
        
        # Images
        self.current_image = None
        self.highlighted_image = None
        self.page_offset_x = 0
        self.page_offset_y = 0
        self.page_images = []  # Voor continuous scroll
        self.page_pil_images = {}  # PIL images voor elke pagina (voor highlighting)
        self.page_positions = []  # Y-positie van elke pagina
        self.scroll_to_page = None  # Flag voor initiële scroll
        
        # Form fields
        self.form_widgets = []
        self.form_data = {}  # Store form field values

    def close_document(self):
        if self.pdf_document:
            self.pdf_document.close()
            self.pdf_document = None

class NVictReader:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("NVict Reader")
        self.root.geometry("1366x768")
        self.root.minsize(800, 600)

        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

        self.config = {"theme": "Systeemstandaard"}
        
        # Update instellingen laden
        self.load_update_settings()
        
        # Herstel opgeslagen schermgrootte en positie
        if self.update_settings.get('window_geometry'):
            try:
                self.root.geometry(self.update_settings['window_geometry'])
            except:
                self.root.geometry("1366x768")  # Fallback naar standaard
        
        # Herstel window state (maximized of normal)
        if self.update_settings.get('window_state') == 'zoomed':
            self.root.state('zoomed')
        
        self.apply_theme()
        
        self.load_icons()
        
        self.setup_ui()
        
        self.setup_shortcuts()
        
        self.update_ui_state()
        
        # Initialiseer drag-and-drop ondersteuning
        self.setup_drag_and_drop()
        
        # Check first run en vraag om default PDF viewer te worden
        if self.update_settings.get('first_run', True):
            self.root.after(1000, self.check_first_run)

    def check_first_run(self):
        """Check first run and show welcome dialog"""
        # Check eerst of we al de standaard PDF viewer zijn
        if DefaultPDFHandler.is_default_pdf_handler():
            # We zijn al de standaard, geen dialoog tonen
            self.update_settings['first_run'] = False
            self.update_settings['ask_default'] = False
            self.save_update_settings()
            return
        
        # We zijn nog niet de standaard, vraag of gebruiker dit wil instellen
        if self.update_settings.get('ask_default', True):
            result = DefaultPDFHandler.show_first_run_dialog(self.root)
            if result == "never":
                self.update_settings['ask_default'] = False
        
        self.update_settings['first_run'] = False
        self.save_update_settings()

    def setup_drag_and_drop(self):
        """Configureer drag-and-drop ondersteuning voor PDF bestanden"""
        # Probeer tkinterdnd2 te gebruiken
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # Re-initialiseer root als TkinterDnD root
            # Dit werkt niet na het feit dat root al is gemaakt
            # Daarom geven we instructies aan de gebruiker
            raise ImportError("tkinterdnd2 moet bij installatie worden ingesteld")
            
        except ImportError:
            # tkinterdnd2 is niet beschikbaar
            # Toon instructie aan gebruiker voor alternatieve methode
            print("Drag-and-drop vereist tkinterdnd2 library")
            print("Alternatief: Gebruik rechtsklik -> Openen met -> NVict Reader in Windows Verkenner")
            print("Of gebruik Bestand -> Openen binnen het programma")
            
            # Zorg dat file associations werken (via sys.argv - dit werkt al)
            pass

    def apply_theme(self):
        theme_choice = self.config.get("theme", "Systeemstandaard")
        theme_name = self.get_windows_theme() if theme_choice == "Systeemstandaard" else theme_choice
        self.theme = Theme.LIGHT if theme_name == "Licht" else Theme.DARK
        self.root.configure(bg=self.theme["BG_PRIMARY"])
        
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure("TNotebook", background=self.theme["BG_PRIMARY"], borderwidth=0)
        style.configure("TNotebook.Tab", background=self.theme["BG_PRIMARY"], 
                       foreground=self.theme["TEXT_SECONDARY"], borderwidth=0, padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", self.theme["BG_SECONDARY"])], 
                 foreground=[("selected", self.theme["TEXT_PRIMARY"])])
        style.configure("TScrollbar", background=self.theme["BG_PRIMARY"], 
                       troughcolor=self.theme["BG_SECONDARY"], bordercolor=self.theme["BG_PRIMARY"], 
                       arrowcolor=self.theme["TEXT_PRIMARY"])
        style.map("TScrollbar", background=[('active', self.theme["ACCENT_COLOR"])])

    def get_windows_theme(self):
        try:
            if winreg:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return "Licht" if value == 1 else "Donker"
        except (FileNotFoundError, AttributeError):
            pass
        return "Licht"
        
    def load_icons(self):
        self.icons = {}
        is_dark_theme = self.theme == Theme.DARK
        icon_files = {
            "open": "open.png", "close": "close.png", "print": "print.png", 
            "zoom-in": "zoom-in.png", "zoom-out": "zoom-out.png",
            "reset": "reset.png", "prev-page": "prev-page.png", "next-page": "next-page.png", 
            "first-page": "first-page.png", "last-page": "last-page.png", "info": "info.png", 
            "copy": "copy.png", "search": "search.png", "pdf": "pdf.png",
            "fit-width": "fit-width.png", "save": "save.png", "toolbox": "toolbox.png"
        }
        
        icons_found = 0
        for name, filename in icon_files.items():
            try:
                path = get_resource_path(os.path.join('icons', filename))
                if os.path.exists(path):
                    image = Image.open(path).resize((18, 18), Image.Resampling.LANCZOS)
                    if is_dark_theme:
                        if image.mode == 'RGBA': 
                            r, g, b, a = image.split()
                            rgb_image = Image.merge('RGB', (r, g, b))
                            inverted_image = ImageOps.invert(rgb_image)
                            r2, g2, b2 = inverted_image.split()
                            image = Image.merge('RGBA', (r2, g2, b2, a))
                        else: 
                            image = ImageOps.invert(image)
                    self.icons[name] = ImageTk.PhotoImage(image)
                    icons_found += 1
                else:
                    self.icons[name] = None
                    
            except Exception: 
                self.icons[name] = None

    def setup_ui(self):
        # Maak menubar
        self.create_menubar()
        
        self.create_modern_toolbar()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        self.welcome_frame = tk.Frame(self.notebook, bg=self.theme["BG_PRIMARY"])
        
        # Laad en toon logo
        try:
            logo_path = get_resource_path('Logo.png')
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                # Resize logo als het te groot is (max 180x180)
                logo_image.thumbnail((180, 180), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(self.welcome_frame, image=self.logo_photo, 
                                     bg=self.theme["BG_PRIMARY"])
                logo_label.place(relx=0.5, rely=0.35, anchor="center")
        except Exception:
            pass
        
        welcome_text = "Welkom bij NVict Reader\n\nKlik op 'Openen' of druk op Ctrl+O om een PDF te laden."
        self.welcome_label = tk.Label(self.welcome_frame, text=welcome_text, 
                                      font=Theme.FONT_HEADING, fg=self.theme["TEXT_SECONDARY"], 
                                      bg=self.theme["BG_PRIMARY"], justify=tk.CENTER)
        self.welcome_label.place(relx=0.5, rely=0.55, anchor="center")
        
        self.create_status_bar()

    def create_menubar(self):
        """Maak menubar met bewerken opties"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Bestand menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bestand", menu=file_menu)
        file_menu.add_command(label="Openen...", command=self.open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="Opslaan...", command=self.save_form_data, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Afdrukken...", command=self.print_pdf, accelerator="Ctrl+P")
        file_menu.add_separator()
        file_menu.add_command(label="Sluiten", command=self.close_active_tab, accelerator="Ctrl+W")
        file_menu.add_command(label="Afsluiten", command=self.exit_application, accelerator="Ctrl+Q")
        
        # Bewerken menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bewerken", menu=edit_menu)
        edit_menu.add_command(label="Kopieer tekst", command=self.copy_text, accelerator="Ctrl+C")
        edit_menu.add_command(label="Zoeken...", command=self.show_search_dialog, accelerator="Ctrl+F")
        edit_menu.add_separator()
        edit_menu.add_command(label="Pagina's exporteren...", command=self.export_pages)
        edit_menu.add_command(label="PDF's samenvoegen...", command=self.merge_pdfs)
        edit_menu.add_command(label="Pagina's roteren...", command=self.rotate_pages)
        
        # Beeld menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Beeld", menu=view_menu)
        view_menu.add_command(label="Zoom in", command=self.zoom_in, accelerator="Ctrl++")
        view_menu.add_command(label="Zoom uit", command=self.zoom_out, accelerator="Ctrl+-")
        view_menu.add_command(label="Pasbreedte", command=lambda: self.set_zoom_mode("fit_width"))
        view_menu.add_separator()
        view_menu.add_command(label="Eerste pagina", command=self.first_page)
        view_menu.add_command(label="Vorige pagina", command=self.prev_page, accelerator="←")
        view_menu.add_command(label="Volgende pagina", command=self.next_page, accelerator="→")
        view_menu.add_command(label="Laatste pagina", command=self.last_page)
        
        # Instellingen menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Instellingen", menu=settings_menu)
        settings_menu.add_command(label="Instellen als standaard PDF viewer", command=self.set_as_default_pdf)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="PDF Info", command=self.show_pdf_info)
        help_menu.add_separator()
        help_menu.add_command(label="Controleer op updates...", command=lambda: self.check_for_updates(silent=False))
        help_menu.add_separator()
        help_menu.add_command(label="Over NVict Reader", command=self.show_about)

    def create_modern_toolbar(self):
        toolbar_frame = tk.Frame(self.root, bg=self.theme["BG_SECONDARY"], height=60, 
                                highlightbackground="#e0e0e0", highlightthickness=1)
        toolbar_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(20, 0))
        toolbar_frame.pack_propagate(False)
        
        # Openen, Sluiten, Opslaan, Zoeken
        self.open_btn = self.create_toolbar_button(toolbar_frame, " Openen", "open", 
                                                   self.open_pdf, self.theme["ACCENT_COLOR"])
        self.close_btn = self.create_toolbar_button(toolbar_frame, " Sluiten", "close", 
                                                    self.close_active_tab, self.theme["BG_SECONDARY"])
        self.save_btn = self.create_toolbar_button(toolbar_frame, "", "save", 
                                                   self.save_form_data, self.theme["BG_SECONDARY"])
        self.search_btn = self.create_toolbar_button(toolbar_frame, "", "search", 
                                                     self.show_search_dialog, self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(toolbar_frame)
        
        # Kopiëren, Printen, Info, Bewerken
        self.copy_btn = self.create_toolbar_button(toolbar_frame, "", "copy", 
                                                   self.copy_text, self.theme["BG_SECONDARY"])
        self.print_btn = self.create_toolbar_button(toolbar_frame, "", "print", 
                                                    self.print_pdf, self.theme["BG_SECONDARY"])
        self.info_btn = self.create_toolbar_button(toolbar_frame, "", "info", 
                                                   self.show_pdf_info, self.theme["BG_SECONDARY"])
        self.edit_btn = self.create_toolbar_button(toolbar_frame, " Bewerken", "toolbox", 
                                                   self.show_edit_menu, self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(toolbar_frame)
        
        # Zoom knoppen
        self.zoom_in_btn = self.create_toolbar_button(toolbar_frame, "", "zoom-in", 
                                                      self.zoom_in, self.theme["BG_SECONDARY"])
        self.zoom_out_btn = self.create_toolbar_button(toolbar_frame, "", "zoom-out", 
                                                       self.zoom_out, self.theme["BG_SECONDARY"])
        self.fit_width_btn = self.create_toolbar_button(toolbar_frame, "", "fit-width", 
                                                        lambda: self.set_zoom_mode("fit_width"), 
                                                        self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(toolbar_frame)
        
        # Pagina navigatie
        self.prev_btn = self.create_toolbar_button(toolbar_frame, "", "prev-page", 
                                                   self.prev_page, self.theme["BG_SECONDARY"])
        self.next_btn = self.create_toolbar_button(toolbar_frame, "", "next-page", 
                                                   self.next_page, self.theme["BG_SECONDARY"])
        
        page_frame = tk.Frame(toolbar_frame, bg=self.theme["BG_SECONDARY"])
        page_frame.pack(side=tk.LEFT, padx=5)
        
        self.page_var = tk.StringVar(value="1")
        page_entry = tk.Entry(page_frame, textvariable=self.page_var, width=5, 
                             font=Theme.FONT_MAIN, justify=tk.CENTER,
                             bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], 
                             relief="flat", bd=1)
        page_entry.pack(side=tk.LEFT)
        page_entry.bind("<Return>", self.go_to_page)
        
        self.total_pages_label = tk.Label(page_frame, text="/ 0", font=Theme.FONT_MAIN, 
                                          fg=self.theme["TEXT_SECONDARY"], 
                                          bg=self.theme["BG_SECONDARY"])
        self.total_pages_label.pack(side=tk.LEFT, padx=(5, 0))

    def create_toolbar_button(self, parent, text, icon_name, command, bg_color):
        btn_frame = tk.Frame(parent, bg=bg_color, highlightthickness=0)
        btn_frame.pack(side=tk.LEFT, padx=5, pady=10)
        
        # Controleer of icon bestaat
        icon_image = self.icons.get(icon_name)
        
        # Emoji fallbacks voor ontbrekende iconen
        emoji_fallbacks = {
            "open": "📂",
            "close": "✕",
            "save": "💾",
            "print": "🖨️",
            "copy": "📋",
            "search": "🔍",
            "zoom-in": "🔍+",
            "zoom-out": "🔍-",
            "reset": "↺",
            "fit-width": "⬌",
            "prev-page": "◄",
            "next-page": "►",
            "first-page": "⏮",
            "last-page": "⏭",
            "info": "ℹ️",
            "toolbox": "🛠️"
        }
        
        # Als icon bestaat, gebruik icon + tekst
        if icon_image:
            btn = tk.Button(btn_frame, text=text, image=icon_image, 
                           compound=tk.LEFT, command=command, font=Theme.FONT_MAIN,
                           bg=bg_color, fg=self.theme["TEXT_PRIMARY"], 
                           activebackground=self.theme["ACCENT_COLOR"],
                           activeforeground="#ffffff", relief="flat", bd=0, 
                           padx=10, pady=5, cursor="hand2")
        else:
            # Geen icon - gebruik emoji fallback
            display_text = text
            if not text and icon_name in emoji_fallbacks:
                display_text = emoji_fallbacks[icon_name]
            elif text and icon_name in emoji_fallbacks:
                # Voeg emoji toe voor de tekst
                display_text = emoji_fallbacks[icon_name] + text
            
            btn = tk.Button(btn_frame, text=display_text, command=command, font=Theme.FONT_MAIN,
                           bg=bg_color, fg=self.theme["TEXT_PRIMARY"], 
                           activebackground=self.theme["ACCENT_COLOR"],
                           activeforeground="#ffffff", relief="flat", bd=0, 
                           padx=10, pady=5, cursor="hand2")
        
        btn.pack()
        
        def on_enter(e):
            btn.configure(bg=self.theme["ACCENT_COLOR"], fg="#ffffff")
        def on_leave(e):
            btn.configure(bg=bg_color, fg=self.theme["TEXT_PRIMARY"])
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn

    def add_toolbar_separator(self, parent):
        separator = tk.Frame(parent, bg=self.theme["TEXT_SECONDARY"], width=1, height=40)
        separator.pack(side=tk.LEFT, padx=10, pady=10)

    def create_status_bar(self):
        status_bar = tk.Frame(self.root, bg=self.theme["BG_SECONDARY"], height=30, 
                             highlightbackground="#e0e0e0", highlightthickness=1)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 20))
        status_bar.pack_propagate(False)
        
        self.status_label = tk.Label(status_bar, text="Klaar", font=Theme.FONT_SMALL, 
                                     fg=self.theme["TEXT_SECONDARY"], bg=self.theme["BG_SECONDARY"], 
                                     anchor="w")
        self.status_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        copyright_label = tk.Label(status_bar, text=f"© {self.get_current_year()} NVict Service - www.nvict.nl", 
                                   font=Theme.FONT_SMALL, fg=self.theme["TEXT_SECONDARY"], 
                                   bg=self.theme["BG_SECONDARY"], anchor="e", cursor="hand2")
        copyright_label.pack(side=tk.RIGHT, padx=10)
        copyright_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.nvict.nl/software.html"))

    def get_current_year(self):
        """Get current year"""
        from datetime import datetime
        return datetime.now().year

    def setup_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.open_pdf())
        self.root.bind("<Control-s>", lambda e: self.save_form_data())
        self.root.bind("<Control-p>", lambda e: self.print_pdf())
        self.root.bind("<Control-w>", lambda e: self.close_active_tab())
        self.root.bind("<Control-q>", lambda e: self.exit_application())
        self.root.bind("<Control-plus>", lambda e: self.zoom_in())
        self.root.bind("<Control-minus>", lambda e: self.zoom_out())
        self.root.bind("<Control-c>", lambda e: self.copy_text())
        self.root.bind("<Control-f>", lambda e: self.show_search_dialog())
        self.root.bind("<Left>", lambda e: self.prev_page())
        self.root.bind("<Right>", lambda e: self.next_page())

    def load_update_settings(self):
        """Laad update instellingen van bestand"""
        self.update_settings = {
            'auto_check': True,  # Automatisch controleren bij opstarten
            'auto_download': False,  # Automatisch downloaden (standaard uit)
            'last_check': None,
            'window_geometry': None,  # Laatst gebruikte schermgrootte
            'window_state': 'normal'  # normal of zoomed (maximized)
        }
        
        try:
            settings_path = get_settings_path()
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    saved_settings = json.load(f)
                    self.update_settings.update(saved_settings)
        except Exception:
            pass  # Gebruik default settings
    
    def save_update_settings(self):
        """Sla update instellingen op naar bestand"""
        try:
            settings_path = get_settings_path()
            with open(settings_path, 'w') as f:
                json.dump(self.update_settings, f, indent=2)
        except Exception:
            pass  # Stille fout
    
    def check_for_updates_on_startup(self):
        """Controleer automatisch op updates bij opstarten (in achtergrond)"""
        if not self.update_settings.get('auto_check', True):
            return
        
        # Doe check in aparte thread om UI niet te blokkeren
        def background_check():
            import time
            time.sleep(2)  # Wacht 2 seconden na opstarten
            self.root.after(0, lambda: self.check_for_updates(silent=True))
        
        thread = threading.Thread(target=background_check, daemon=True)
        thread.start()

    def get_active_tab(self):
        try:
            current_tab_id = self.notebook.select()
            if current_tab_id:
                return self.notebook.nametowidget(current_tab_id)
        except:
            pass
        return None

    def on_tab_change(self, event=None):
        self.update_ui_state()

    def update_ui_state(self):
        tab = self.get_active_tab()
        has_pdf = isinstance(tab, PDFTab)
        
        for btn in [self.close_btn, self.save_btn, self.print_btn, self.zoom_in_btn, 
                   self.zoom_out_btn, self.fit_width_btn, self.prev_btn, 
                   self.next_btn, self.copy_btn, self.search_btn, self.info_btn, self.edit_btn]:
            btn.config(state=tk.NORMAL if has_pdf else tk.DISABLED)
        
        if has_pdf:
            self.page_var.set(str(tab.current_page + 1))
            self.total_pages_label.config(text=f"/ {len(tab.pdf_document)}")
            self.status_label.config(text=f"Zoom: {int(tab.zoom_level * 100)}%")
        else:
            self.page_var.set("1")
            self.total_pages_label.config(text="/ 0")
            self.status_label.config(text="Geen document geopend")

    def open_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Selecteer een PDF bestand",
            filetypes=[("PDF Bestanden", "*.pdf"), ("Alle Bestanden", "*.*")]
        )
        if file_path:
            self.add_new_tab(file_path)

    def add_new_tab(self, file_path):
        try:
            # Probeer eerst te openen om te controleren of wachtwoord nodig is
            test_doc = fitz.open(file_path)
            
            # Controleer of document beveiligd is
            if test_doc.needs_pass:
                test_doc.close()
                
                # Vraag wachtwoord
                password = self.ask_password(file_path)
                
                if password is None:
                    # Gebruiker heeft geannuleerd
                    return
                
                # Probeer te openen met wachtwoord
                test_doc = fitz.open(file_path)
                auth_result = test_doc.authenticate(password)
                
                if not auth_result:
                    test_doc.close()
                    messagebox.showerror("Fout", 
                        "Onjuist wachtwoord!\n\nKan het PDF bestand niet openen.")
                    return
                
                # Wachtwoord is correct, sla het op voor later gebruik
                # (wordt gebruikt bij PDFTab aanmaak)
                self.temp_password = password
            else:
                self.temp_password = None
            
            test_doc.close()
            
            # Nu de tab aanmaken
            if self.welcome_frame.winfo_ismapped():
                self.notebook.forget(self.welcome_frame)

            tab = PDFTab(self.notebook, file_path, self.theme, getattr(self, 'temp_password', None))
            self.notebook.add(tab, text=os.path.basename(file_path), padding=5)
            self.notebook.select(tab)
            self.display_page(tab)
            
            # Bind events
            tab.canvas.bind("<Configure>", lambda e, t=tab: self.on_resize(e, t))
            tab.canvas.bind("<Button-1>", lambda e, t=tab: self.on_click(e, t))
            tab.canvas.bind("<B1-Motion>", lambda e, t=tab: self.on_drag(e, t))
            tab.canvas.bind("<ButtonRelease-1>", lambda e, t=tab: self.on_release(e, t))
            
            # Muiswiel
            tab.canvas.bind("<MouseWheel>", lambda e, t=tab: self.on_mousewheel(e, t))
            tab.canvas.bind("<Button-4>", lambda e, t=tab: self.on_mousewheel(e, t))
            tab.canvas.bind("<Button-5>", lambda e, t=tab: self.on_mousewheel(e, t))
            
            # Wis tijdelijk wachtwoord
            self.temp_password = None
            
        except Exception as e:
            messagebox.showerror("Fout", f"Kan PDF niet openen:\n{str(e)}")
    
    def ask_password(self, file_path):
        """Toon dialoog om wachtwoord op te vragen voor beveiligde PDF"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Wachtwoord vereist")
        dialog.geometry("450x280")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur
        header_frame = tk.Frame(dialog, bg=self.theme["WARNING_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔒 Wachtwoord Vereist", font=("Segoe UI", 14, "bold"),
                bg=self.theme["WARNING_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        tk.Label(content_frame, 
                text=f"Dit PDF bestand is beveiligd met een wachtwoord.\n\nBestand: {os.path.basename(file_path)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_PRIMARY"],
                justify=tk.LEFT,
                wraplength=380).pack(pady=(0, 20))
        
        tk.Label(content_frame, text="Wachtwoord:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        password_var = tk.StringVar()
        password_entry = tk.Entry(content_frame, textvariable=password_var, 
                                 font=("Segoe UI", 10), show="●", width=40)
        password_entry.pack(pady=8, fill=tk.X)
        password_entry.focus()
        
        result = {"password": None}
        
        def on_ok():
            result["password"] = password_var.get()
            dialog.destroy()
        
        def on_cancel():
            result["password"] = None
            dialog.destroy()
        
        # Footer met knoppen
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="OK", command=on_ok,
                 bg=self.theme["WARNING_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=on_cancel,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Enter key binding
        password_entry.bind("<Return>", lambda e: on_ok())
        
        # Wacht tot dialoog gesloten is
        self.root.wait_window(dialog)
        
        return result["password"]

    def on_mousewheel(self, event, tab):
        if event.num == 4 or event.delta > 0:
            tab.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            tab.canvas.yview_scroll(1, "units")

    def close_active_tab(self):
        active_tab = self.get_active_tab()
        if isinstance(active_tab, PDFTab):
            active_tab.close_document()
            self.notebook.forget(active_tab)
            if len(self.notebook.tabs()) == 0:
                self.notebook.add(self.welcome_frame)
            self.update_ui_state()
    
    def on_resize(self, event, tab):
        if tab.zoom_mode == "fit_width":
            self.display_page(tab)

    def display_page(self, tab):
        if not tab or not tab.pdf_document:
            return

        # Clear
        tab.canvas.delete("all")
        tab.text_words = []
        tab.selected_text = ""
        
        # Verwijder oude form widgets
        for widget in tab.form_widgets:
            widget.destroy()
        tab.form_widgets = []

        # Bereken zoom voor fit_width mode
        if tab.zoom_mode == "fit_width":
            canvas_width = tab.canvas.winfo_width() - 40
            page_width = tab.pdf_document[0].bound().width
            if page_width > 0:
                tab.zoom_level = canvas_width / page_width
        
        # Render ALLE pagina's onder elkaar
        x_offset = 20  # Links marge
        y_offset = 20  # Begin positie
        page_spacing = 20  # Ruimte tussen pagina's
        
        # Houd bij waar elke pagina begint
        tab.page_positions = []
        
        for page_num in range(len(tab.pdf_document)):
            page = tab.pdf_document[page_num]
            
            # Sla positie op
            tab.page_positions.append(y_offset)
            
            # Render pagina
            mat = fitz.Matrix(tab.zoom_level, tab.zoom_level)
            pix = page.get_pixmap(matrix=mat)
            
            img_data = pix.tobytes("ppm")
            pil_image = Image.open(io.BytesIO(img_data))
            
            # Bewaar afbeelding (voor highlights later)
            if page_num == tab.current_page:
                tab.current_image = pil_image.copy()
            
            # Bewaar alle pagina afbeeldingen voor selectie highlighting
            if not hasattr(tab, 'page_pil_images'):
                tab.page_pil_images = {}
            tab.page_pil_images[page_num] = pil_image.copy()
            
            photo = ImageTk.PhotoImage(pil_image)
            
            # Sla referentie op zodat garbage collector het niet verwijdert
            if not hasattr(tab, 'page_images'):
                tab.page_images = []
            if len(tab.page_images) <= page_num:
                tab.page_images.extend([None] * (page_num - len(tab.page_images) + 1))
            tab.page_images[page_num] = photo
            
            # Teken pagina op canvas
            img_width, img_height = pil_image.size
            tab.canvas.create_image(x_offset, y_offset, anchor="nw", image=photo, tags=f"page_{page_num}")
            
            # Teken pagina nummer
            page_num_text = f"Pagina {page_num + 1} / {len(tab.pdf_document)}"
            tab.canvas.create_text(
                x_offset + img_width // 2, 
                y_offset - 5,
                text=page_num_text,
                font=("Segoe UI", 9),
                fill=self.theme["TEXT_SECONDARY"]
            )
            
            # Extract tekst voor deze pagina
            words = page.get_text("words")
            for word_info in words:
                x0, y0, x1, y1, text = word_info[0], word_info[1], word_info[2], word_info[3], word_info[4]
                sx0 = x0 * tab.zoom_level + x_offset
                sy0 = y0 * tab.zoom_level + y_offset
                sx1 = x1 * tab.zoom_level + x_offset
                sy1 = y1 * tab.zoom_level + y_offset
                tab.text_words.append((text, sx0, sy0, sx1, sy1))
            
            # Toon formuliervelden voor deze pagina
            self.display_form_fields_for_page(tab, page, page_num, x_offset, y_offset)
            
            # Teken lichte lijn onder pagina (scheiding)
            separator_y = y_offset + img_height + page_spacing // 2
            tab.canvas.create_line(
                x_offset, separator_y,
                x_offset + img_width, separator_y,
                fill=self.theme["TEXT_SECONDARY"], width=1, dash=(2, 4)
            )
            
            # Update offset voor volgende pagina
            y_offset += img_height + page_spacing
        
        # Sla offset info op voor navigatie
        tab.page_offset_x = x_offset
        tab.page_offset_y = 20
        
        # Als we naar een specifieke pagina navigeren, scroll erheen
        if hasattr(tab, 'scroll_to_page') and tab.scroll_to_page is not None:
            self.scroll_to_page(tab, tab.scroll_to_page)
            tab.scroll_to_page = None
        
        # Scrollregion instellen
        total_height = y_offset + 20
        max_width = max([tab.pdf_document[i].bound().width * tab.zoom_level for i in range(len(tab.pdf_document))]) + x_offset * 2
        tab.canvas.configure(scrollregion=(0, 0, max_width, total_height))
        
        self.update_ui_state()

    def scroll_to_page(self, tab, page_num):
        """Scroll naar een specifieke pagina"""
        if not hasattr(tab, 'page_positions') or page_num >= len(tab.page_positions):
            return
        
        y_pos = tab.page_positions[page_num]
        
        # Scroll canvas naar deze positie
        total_height = int(tab.canvas.cget("scrollregion").split()[3])
        canvas_height = tab.canvas.winfo_height()
        
        if total_height > canvas_height:
            # Bereken fractie voor scrollpositie (0.0 - 1.0)
            fraction = y_pos / total_height
            tab.canvas.yview_moveto(fraction)

    def display_form_fields_for_page(self, tab, page, page_num, x_offset, y_offset):
        """Toon formuliervelden voor een specifieke pagina"""
        try:
            widgets = list(page.widgets())
            
            if not widgets:
                return
            
            for widget in widgets:
                field_type = widget.field_type
                field_name = widget.field_name
                rect = widget.rect
                
                # Schaal en positioneer het veld (relatief aan pagina positie)
                x0 = rect.x0 * tab.zoom_level + x_offset
                y0 = rect.y0 * tab.zoom_level + y_offset
                x1 = rect.x1 * tab.zoom_level + x_offset
                y1 = rect.y1 * tab.zoom_level + y_offset
                
                width = x1 - x0
                height = y1 - y0
                
                # Maak field widgets (zelfde als voorheen)
                if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    current_value = widget.field_value or ""
                    if field_name in tab.form_data:
                        current_value = tab.form_data[field_name]
                    
                    var = tk.StringVar(value=current_value)
                    entry = tk.Entry(tab.canvas, textvariable=var,
                                    font=("Arial", max(8, int(height * 0.6))),
                                    bg="white", fg="black", relief="solid", bd=1)
                    
                    def save_value(name=field_name, variable=var, w=widget):
                        tab.form_data[name] = variable.get()
                        w.field_value = variable.get()
                        w.update()
                    
                    var.trace('w', lambda *args, sv=save_value: sv())
                    
                    window = tab.canvas.create_window(x0, y0, anchor="nw", 
                                                     window=entry, width=width, height=height)
                    tab.form_widgets.append(entry)
                
                elif field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    current_value = widget.field_value
                    if field_name in tab.form_data:
                        current_value = tab.form_data[field_name]
                    
                    var = tk.BooleanVar(value=bool(current_value))
                    
                    checkbox = tk.Checkbutton(tab.canvas, variable=var,
                                             bg=self.theme["BG_PRIMARY"],
                                             activebackground=self.theme["BG_PRIMARY"],
                                             selectcolor="white")
                    
                    def save_checkbox(name=field_name, variable=var, w=widget):
                        tab.form_data[name] = variable.get()
                        w.field_value = variable.get()
                        w.update()
                    
                    var.trace('w', lambda *args, sc=save_checkbox: sc())
                    
                    window = tab.canvas.create_window(x0, y0, anchor="nw",
                                                     window=checkbox, width=width, height=height)
                    tab.form_widgets.append(checkbox)
                
                elif field_type == fitz.PDF_WIDGET_TYPE_COMBOBOX:
                    current_value = widget.field_value or ""
                    if field_name in tab.form_data:
                        current_value = tab.form_data[field_name]
                    
                    choices = widget.choice_values if hasattr(widget, 'choice_values') else []
                    
                    var = tk.StringVar(value=current_value)
                    combo = ttk.Combobox(tab.canvas, textvariable=var, values=choices,
                                        font=("Arial", max(8, int(height * 0.6))),
                                        state="readonly")
                    
                    def save_combo(name=field_name, variable=var, w=widget):
                        tab.form_data[name] = variable.get()
                        w.field_value = variable.get()
                        w.update()
                    
                    var.trace('w', lambda *args, sco=save_combo: sco())
                    
                    window = tab.canvas.create_window(x0, y0, anchor="nw",
                                                     window=combo, width=width, height=height)
                    tab.form_widgets.append(combo)
                    
        except Exception as e:
            print(f"Fout bij tonen formuliervelden op pagina {page_num + 1}: {e}")

    def on_click(self, event, tab):
        # Clear oude selectie
        tab.selected_text = ""
        
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        tab.drag_start = (x, y)
        
        # Herstel alle originele pagina afbeeldingen (verwijder highlighting)
        if hasattr(tab, 'page_pil_images') and hasattr(tab, 'page_positions'):
            for page_num in tab.page_pil_images.keys():
                original_image = tab.page_pil_images[page_num]
                photo = ImageTk.PhotoImage(original_image)
                
                # Bewaar photo reference
                if len(tab.page_images) > page_num:
                    tab.page_images[page_num] = photo
                
                # Update de afbeelding op canvas
                page_y_offset = tab.page_positions[page_num]
                page_x_offset = tab.page_offset_x
                
                tab.canvas.delete(f"page_{page_num}")
                tab.canvas.create_image(page_x_offset, page_y_offset, 
                                       anchor="nw", image=photo, tags=f"page_{page_num}")
        
        # Teken drag rectangle
        if tab.drag_rect:
            tab.canvas.delete(tab.drag_rect)
        tab.drag_rect = tab.canvas.create_rectangle(
            x, y, x, y, outline="red", width=2, tags="drag_rect"
        )
        
        self.status_label.config(text="Selecteren...")

    def on_drag(self, event, tab):
        if not tab.drag_start:
            return
        
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        
        if tab.drag_rect:
            tab.canvas.coords(tab.drag_rect, tab.drag_start[0], tab.drag_start[1], x, y)

    def on_release(self, event, tab):
        if not tab.drag_start:
            return
        
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        
        # Verwijder drag rectangle
        if tab.drag_rect:
            tab.canvas.delete(tab.drag_rect)
            tab.drag_rect = None
        
        # Selection bounds
        x1, y1 = tab.drag_start
        x2, y2 = x, y
        
        left = min(x1, x2)
        right = max(x1, x2)
        top = min(y1, y2)
        bottom = max(y1, y2)
        
        # Detecteer op welke pagina(s) de selectie zich bevindt
        selected_pages = set()
        for word_data in tab.text_words:
            text, wx0, wy0, wx1, wy1 = word_data
            if not (wx1 < left or wx0 > right or wy1 < top or wy0 > bottom):
                # Bepaal pagina nummer op basis van Y-positie
                for page_num, page_y_pos in enumerate(tab.page_positions):
                    if page_num + 1 < len(tab.page_positions):
                        next_page_y = tab.page_positions[page_num + 1]
                        if page_y_pos <= wy0 < next_page_y:
                            selected_pages.add(page_num)
                            break
                    else:
                        # Laatste pagina
                        if wy0 >= page_y_pos:
                            selected_pages.add(page_num)
                            break
        
        # Vind geselecteerde woorden
        selected_words = []
        for word_data in tab.text_words:
            text, wx0, wy0, wx1, wy1 = word_data
            if not (wx1 < left or wx0 > right or wy1 < top or wy0 > bottom):
                selected_words.append(word_data)
        
        if len(selected_words) > 0:
            # Sorteer woorden
            selected_words.sort(key=lambda w: (w[2], w[1]))
            
            # Voor elke geselecteerde pagina, maak een highlighted versie
            for page_num in selected_pages:
                if not hasattr(tab, 'page_pil_images') or page_num not in tab.page_pil_images:
                    continue
                
                # Haal de originele afbeelding van deze pagina op
                original_image = tab.page_pil_images[page_num].copy()
                highlighted = original_image.copy()
                draw = ImageDraw.Draw(highlighted, 'RGBA')
                
                # Bepaal offset voor deze pagina
                page_y_offset = tab.page_positions[page_num]
                page_x_offset = tab.page_offset_x
                
                # Teken highlights alleen voor woorden op deze pagina
                for word_data in selected_words:
                    text, wx0, wy0, wx1, wy1 = word_data
                    
                    # Check of dit woord op deze pagina is
                    word_on_this_page = False
                    if page_num + 1 < len(tab.page_positions):
                        next_page_y = tab.page_positions[page_num + 1]
                        if page_y_offset <= wy0 < next_page_y:
                            word_on_this_page = True
                    else:
                        if wy0 >= page_y_offset:
                            word_on_this_page = True
                    
                    if not word_on_this_page:
                        continue
                    
                    # Bereken positie relatief aan deze specifieke pagina afbeelding
                    rel_x0 = wx0 - page_x_offset
                    rel_y0 = wy0 - page_y_offset
                    rel_x1 = wx1 - page_x_offset
                    rel_y1 = wy1 - page_y_offset
                    
                    # Teken semi-transparante gele rechthoek
                    draw.rectangle(
                        [rel_x0, rel_y0, rel_x1, rel_y1],
                        fill=(255, 215, 0, 100),  # Goud met 100/255 opacity
                        outline=(255, 165, 0, 200),  # Oranje rand
                        width=2
                    )
                
                # Update canvas met gehighlighte afbeelding voor deze pagina
                photo = ImageTk.PhotoImage(highlighted)
                
                # Bewaar de photo reference
                if len(tab.page_images) > page_num:
                    tab.page_images[page_num] = photo
                
                # Verwijder oude afbeelding en teken nieuwe
                tab.canvas.delete(f"page_{page_num}")
                tab.canvas.create_image(page_x_offset, page_y_offset, 
                                       anchor="nw", image=photo, tags=f"page_{page_num}")
            
            # Verzamel tekst
            tab.selected_text = ""
            last_y = None
            for word_data in selected_words:
                text, wx0, wy0, wx1, wy1 = word_data
                
                if last_y is not None and abs(wy0 - last_y) > 5:
                    tab.selected_text += "\n"
                elif tab.selected_text:
                    tab.selected_text += " "
                
                tab.selected_text += text
                last_y = wy0
            
            self.status_label.config(text=f"Geselecteerd: {len(tab.selected_text)} tekens")
        else:
            self.status_label.config(text="Geen tekst geselecteerd")
        
        tab.drag_start = None

    def copy_text(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab) and tab.selected_text:
            self.root.clipboard_clear()
            self.root.clipboard_append(tab.selected_text)
            self.status_label.config(text=f"Gekopieerd: {len(tab.selected_text)} tekens")
        else:
            self.status_label.config(text="Geen tekst geselecteerd")

    def show_search_dialog(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            search_window = tk.Toplevel(self.root)
            search_window.title("Zoeken in PDF")
            search_window.geometry("450x240")
            search_window.configure(bg=self.theme["BG_PRIMARY"])
            search_window.transient(self.root)
            search_window.grab_set()
            search_window.resizable(False, False)
            
            # Icon toevoegen
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    search_window.iconbitmap(icon_path)
            except:
                pass
            
            # Header met accent kleur (moderne stijl)
            header_frame = tk.Frame(search_window, bg=self.theme["ACCENT_COLOR"], height=60)
            header_frame.pack(fill=tk.X)
            header_frame.pack_propagate(False)
            
            tk.Label(header_frame, text="🔍 Zoeken in PDF", font=("Segoe UI", 14, "bold"),
                    bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
            
            # Content frame
            content_frame = tk.Frame(search_window, bg=self.theme["BG_PRIMARY"])
            content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
            
            tk.Label(content_frame, text="Zoek tekst:", font=Theme.FONT_MAIN,
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(pady=(0, 10))
            
            search_var = tk.StringVar()
            search_entry = tk.Entry(content_frame, textvariable=search_var, 
                                   font=Theme.FONT_MAIN, width=40)
            search_entry.pack(pady=5)
            search_entry.focus()
            
            # Footer met knoppen (moderne stijl)
            footer_frame = tk.Frame(search_window, bg=self.theme["BG_SECONDARY"], height=70)
            footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
            footer_frame.pack_propagate(False)
            
            btn_frame = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
            btn_frame.pack(expand=True)
            
            def do_search():
                search_text = search_var.get()
                if search_text:
                    self.search_in_pdf(tab, search_text)
                    search_window.destroy()
            
            tk.Button(btn_frame, text="Zoeken", command=do_search, 
                     bg=self.theme["ACCENT_COLOR"], fg="white", 
                     font=("Segoe UI", 10), padx=25, pady=10,
                     relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
            
            tk.Button(btn_frame, text="Annuleren", command=search_window.destroy,
                     bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                     font=("Segoe UI", 10), padx=25, pady=10,
                     relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
            
            search_entry.bind("<Return>", lambda e: do_search())

    def search_in_pdf(self, tab, search_text):
        found = False
        start_page = tab.current_page
        
        for offset in range(len(tab.pdf_document)):
            page_num = (start_page + offset) % len(tab.pdf_document)
            page = tab.pdf_document[page_num]
            instances = page.search_for(search_text)
            
            if instances:
                if page_num != tab.current_page:
                    tab.current_page = page_num
                    self.display_page(tab)
                
                # Highlight op de afbeelding
                highlighted = tab.current_image.copy()
                draw = ImageDraw.Draw(highlighted, 'RGBA')
                
                for inst in instances:
                    rect = fitz.Rect(inst)
                    x0 = rect.x0 * tab.zoom_level
                    y0 = rect.y0 * tab.zoom_level
                    x1 = rect.x1 * tab.zoom_level
                    y1 = rect.y1 * tab.zoom_level
                    
                    draw.rectangle(
                        [x0, y0, x1, y1],
                        outline=(255, 140, 0, 255),  # Oranje
                        width=3
                    )
                
                photo = ImageTk.PhotoImage(highlighted)
                tab.highlighted_image = photo
                tab.canvas.delete("page_image")
                tab.canvas.create_image(tab.page_offset_x, tab.page_offset_y,
                                       anchor="nw", image=photo, tags="page_image")
                
                found = True
                self.status_label.config(
                    text=f"Gevonden: '{search_text}' ({len(instances)}x op pagina {page_num + 1})"
                )
                break
        
        if not found:
            messagebox.showinfo("Zoeken", f"'{search_text}' niet gevonden in document")

    # Navigatie functies
    def navigate(self, delta):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            new_page = tab.current_page + delta
            if 0 <= new_page < len(tab.pdf_document):
                tab.current_page = new_page
                self.scroll_to_page(tab, tab.current_page)
                self.update_ui_state()

    def first_page(self): 
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.current_page = 0
            self.scroll_to_page(tab, 0)
            self.update_ui_state()

    def prev_page(self): 
        self.navigate(-1)
        
    def next_page(self): 
        self.navigate(1)
    
    def last_page(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.current_page = len(tab.pdf_document) - 1
            self.scroll_to_page(tab, tab.current_page)
            self.update_ui_state()

    def go_to_page(self, event=None):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            try:
                page_num = int(self.page_var.get()) - 1
                if 0 <= page_num < len(tab.pdf_document):
                    tab.current_page = page_num
                    self.scroll_to_page(tab, page_num)
                    self.update_ui_state()
            except ValueError:
                self.update_ui_state()

    # Zoom functies
    def zoom(self, factor):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.zoom_mode = "manual"
            new_zoom = tab.zoom_level * factor
            if 0.2 < new_zoom < 5.0:
                tab.zoom_level = new_zoom
                self.display_page(tab)
    
    def zoom_in(self): 
        self.zoom(1.2)
        
    def zoom_out(self): 
        self.zoom(1/1.2)

    def set_zoom_mode(self, mode):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.zoom_mode = mode
            self.display_page(tab)

    def print_pdf(self):
        """Toon ingebouwde print dialoog met printer selectie"""
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            print_dialog = tk.Toplevel(self.root)
            print_dialog.title("Afdrukken")
            print_dialog.geometry("550x600")
            print_dialog.configure(bg=self.theme["BG_PRIMARY"])
            print_dialog.transient(self.root)
            print_dialog.grab_set()
            print_dialog.resizable(False, False)
            
            # Voeg logo toe aan taakbalk
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    print_dialog.iconbitmap(icon_path)
            except:
                pass
            
            # Header
            header_frame = tk.Frame(print_dialog, bg=self.theme["ACCENT_COLOR"], height=60)
            header_frame.pack(fill=tk.X)
            header_frame.pack_propagate(False)
            
            tk.Label(header_frame, text="🖨️ PDF Afdrukken", font=("Segoe UI", 14, "bold"),
                    bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
            
            # Main content
            content_frame = tk.Frame(print_dialog, bg=self.theme["BG_PRIMARY"])
            content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
            
            # Document info
            tk.Label(content_frame, text="Document:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 2))
            tk.Label(content_frame, text=os.path.basename(tab.file_path), 
                    font=("Segoe UI", 9),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], anchor="w").pack(fill=tk.X, pady=(0, 15))
            
            # Printer selectie
            tk.Label(content_frame, text="Printer:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            printers = self.get_available_printers()
            printer_var = tk.StringVar(value=printers[0] if printers else "Standaard printer")
            
            # Frame voor combobox met border voor betere zichtbaarheid
            combo_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], 
                                  highlightbackground=self.theme["TEXT_SECONDARY"],
                                  highlightthickness=1)
            combo_frame.pack(fill=tk.X, pady=(0, 15))
            
            # Combobox met goede contrast kleuren
            printer_dropdown = ttk.Combobox(combo_frame, textvariable=printer_var, 
                                           values=printers, state="readonly", 
                                           font=("Segoe UI", 10))
            printer_dropdown.pack(fill=tk.X, padx=2, pady=2)
            
            # Override combobox kleuren voor beter contrast
            self.root.option_add('*TCombobox*Listbox.background', 'white')
            self.root.option_add('*TCombobox*Listbox.foreground', 'black')
            self.root.option_add('*TCombobox*Listbox.selectBackground', self.theme["ACCENT_COLOR"])
            self.root.option_add('*TCombobox*Listbox.selectForeground', 'white')
            
            # Pagina selectie
            tk.Label(content_frame, text="Pagina's:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            page_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            page_frame.pack(fill=tk.X, pady=(0, 15))
            
            page_option = tk.StringVar(value="all")
            
            # Alle pagina's
            tk.Radiobutton(page_frame, text=f"Alle pagina's (1-{len(tab.pdf_document)})", 
                          variable=page_option, value="all",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w")
            
            # Huidige pagina
            current_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            current_frame.pack(anchor="w", pady=5)
            tk.Radiobutton(current_frame, text="Huidige pagina", 
                          variable=page_option, value="current",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)
            tk.Label(current_frame, text=f"(pagina {tab.current_page + 1})",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], 
                    font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=5)
            
            # Aangepaste pagina's
            custom_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            custom_frame.pack(anchor="w", pady=5)
            
            tk.Radiobutton(custom_frame, text="Aangepaste pagina's:", 
                          variable=page_option, value="custom",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)
            
            custom_pages_var = tk.StringVar(value="1,3")
            
            # Frame voor entry met border
            entry_frame = tk.Frame(custom_frame, bg=self.theme["BG_SECONDARY"],
                                  highlightbackground=self.theme["TEXT_SECONDARY"],
                                  highlightthickness=1)
            entry_frame.pack(side=tk.LEFT, padx=5)
            
            custom_entry = tk.Entry(entry_frame, textvariable=custom_pages_var,
                                   font=("Segoe UI", 10), width=18,
                                   bg="white",
                                   fg="black",
                                   relief="flat",
                                   insertbackground="black")
            custom_entry.pack(padx=2, pady=2)
            
            # Uitleg voor aangepaste pagina's
            help_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            help_frame.pack(anchor="w", padx=20, pady=2)
            tk.Label(help_frame, text="(bijv: 1,3,5 of 1-3,5 of 2-5)",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], 
                    font=("Segoe UI", 8)).pack(anchor="w")
            
            # Aantal kopieën
            tk.Label(content_frame, text="Aantal kopieën:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            copies_var = tk.StringVar(value="1")
            copies_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            copies_frame.pack(fill=tk.X, pady=(0, 15))
            
            # Frame voor spinbox met border
            spinbox_frame = tk.Frame(copies_frame, bg=self.theme["BG_SECONDARY"],
                                    highlightbackground=self.theme["TEXT_SECONDARY"],
                                    highlightthickness=1)
            spinbox_frame.pack(side=tk.LEFT)
            
            tk.Spinbox(spinbox_frame, from_=1, to=99, textvariable=copies_var,
                      font=("Segoe UI", 10), width=8,
                      bg="white",
                      fg="black",
                      buttonbackground=self.theme["BG_SECONDARY"],
                      relief="flat",
                      insertbackground="black").pack(padx=2, pady=2)
            
            # Passend maken optie
            tk.Label(content_frame, text="Afdruk opties:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            fit_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            fit_frame.pack(fill=tk.X, pady=(0, 10))
            
            fit_to_page_var = tk.BooleanVar(value=True)
            tk.Checkbutton(fit_frame, text="Passend maken op pagina", 
                          variable=fit_to_page_var,
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w")
            tk.Label(fit_frame, text="(schaalt document om op papier te passen)",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], 
                    font=("Segoe UI", 8)).pack(anchor="w", padx=20)
            
            # Knoppen
            btn_frame = tk.Frame(print_dialog, bg=self.theme["BG_SECONDARY"], height=70)
            btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            button_container = tk.Frame(btn_frame, bg=self.theme["BG_SECONDARY"])
            button_container.pack(expand=True)
            
            def do_print():
                try:
                    printer = printer_var.get()
                    copies = int(copies_var.get())
                    page_opt = page_option.get()
                    fit_to_page = fit_to_page_var.get()
                    
                    # Bepaal welke pagina's te printen
                    if page_opt == "current":
                        pages_to_print = [tab.current_page]
                    elif page_opt == "custom":
                        # Parse aangepaste pagina selectie
                        custom_pages = custom_pages_var.get()
                        pages_to_print = self.parse_page_range(custom_pages, len(tab.pdf_document))
                        
                        if not pages_to_print:
                            messagebox.showerror("Ongeldige pagina's", 
                                "Ongeldige pagina selectie.\n\n" +
                                "Gebruik formaat zoals:\n" +
                                "• 1,3,5 (specifieke pagina's)\n" +
                                "• 1-5 (bereik)\n" +
                                "• 1-3,5,7-9 (combinatie)")
                            return
                    else:  # all
                        pages_to_print = list(range(len(tab.pdf_document)))
                    
                    print_dialog.destroy()
                    self.execute_print(tab, printer, pages_to_print, copies, fit_to_page)
                    
                except ValueError as e:
                    messagebox.showerror("Invoer Fout", f"Ongeldig aantal kopieën: {str(e)}")
                except Exception as e:
                    print_dialog.destroy()
                    messagebox.showerror("Print Fout", f"Kan niet printen:\n{str(e)}")
            
            print_btn = tk.Button(button_container, text="Afdrukken", 
                                 command=do_print,
                                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                                 font=("Segoe UI", 10, "bold"), 
                                 padx=30, pady=10,
                                 relief="flat", cursor="hand2")
            print_btn.pack(side=tk.LEFT, padx=5)
            
            cancel_btn = tk.Button(button_container, text="Annuleren", 
                                  command=print_dialog.destroy,
                                  bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                                  font=("Segoe UI", 10), 
                                  padx=30, pady=10,
                                  relief="flat", cursor="hand2")
            cancel_btn.pack(side=tk.LEFT, padx=5)
            
            def on_enter_print(e):
                print_btn.config(bg="#0d8cbd")
            def on_leave_print(e):
                print_btn.config(bg=self.theme["ACCENT_COLOR"])
            def on_enter_cancel(e):
                cancel_btn.config(bg=self.theme["ACCENT_COLOR"], fg="white")
            def on_leave_cancel(e):
                cancel_btn.config(bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"])
            
            print_btn.bind("<Enter>", on_enter_print)
            print_btn.bind("<Leave>", on_leave_print)
            cancel_btn.bind("<Enter>", on_enter_cancel)
            cancel_btn.bind("<Leave>", on_leave_cancel)

    def get_available_printers(self):
        """Haal beschikbare printers op"""
        printers = []
        try:
            if platform.system() == "Windows":
                try:
                    import win32print
                    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
                except ImportError:
                    result = subprocess.run(
                        ["powershell", "-Command", "Get-Printer | Select-Object -ExpandProperty Name"],
                        capture_output=True, text=True, timeout=5, creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if result.returncode == 0:
                        printers = [p.strip() for p in result.stdout.split('\n') if p.strip()]
        except Exception as e:
            print(f"Kon printers niet ophalen: {e}")
        
        if not printers:
            printers = ["Standaard printer"]
        
        return printers

    def parse_page_range(self, page_string, total_pages):
        """Parse pagina bereik string zoals '1,3,5' of '1-5,7' naar lijst van pagina nummers (0-indexed)"""
        pages = set()
        
        try:
            # Verwijder spaties
            page_string = page_string.replace(" ", "")
            
            # Split op komma's
            parts = page_string.split(',')
            
            for part in parts:
                if '-' in part:
                    # Bereik zoals '1-5'
                    start, end = part.split('-')
                    start = int(start)
                    end = int(end)
                    
                    # Valideer bereik
                    if start < 1 or end > total_pages or start > end:
                        return None
                    
                    # Voeg alle pagina's in bereik toe (convert naar 0-indexed)
                    for page_num in range(start, end + 1):
                        pages.add(page_num - 1)
                else:
                    # Enkele pagina zoals '3'
                    page_num = int(part)
                    
                    # Valideer pagina nummer
                    if page_num < 1 or page_num > total_pages:
                        return None
                    
                    # Voeg pagina toe (convert naar 0-indexed)
                    pages.add(page_num - 1)
            
            # Sorteer en return als lijst
            return sorted(list(pages))
            
        except (ValueError, AttributeError):
            return None

    def execute_print(self, tab, printer, pages, copies, fit_to_page=True):
        """Voer print opdracht uit"""
        try:
            temp_path = tempfile.mktemp(suffix=".pdf")
            output_doc = fitz.open()
            
            # Kopieer geselecteerde pagina's
            for page_num in pages:
                output_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
            
            # Als "Passend maken" is aangevinkt, schaal de pagina's
            if fit_to_page:
                # A4 formaat in points (595 x 842)
                page_width = 595
                page_height = 842
                
                for page in output_doc:
                    # Haal huidige pagina afmetingen op
                    rect = page.rect
                    current_width = rect.width
                    current_height = rect.height
                    
                    # Bereken schaalfactor om op A4 te passen (behoud aspect ratio)
                    scale_x = page_width / current_width
                    scale_y = page_height / current_height
                    scale = min(scale_x, scale_y)  # Gebruik kleinste om binnen pagina te blijven
                    
                    # Alleen schalen als de pagina groter is dan A4
                    if scale < 1.0:
                        # Maak transformatie matrix voor schalen
                        mat = fitz.Matrix(scale, scale)
                        
                        # Bereken nieuwe afmetingen
                        new_width = current_width * scale
                        new_height = current_height * scale
                        
                        # Centreer op A4 pagina
                        offset_x = (page_width - new_width) / 2
                        offset_y = (page_height - new_height) / 2
                        
                        # Set nieuwe pagina grootte
                        page.set_mediabox(fitz.Rect(0, 0, page_width, page_height))
                        
                        # Haal alle content op en schaal
                        page.apply_redactions()
            
            output_doc.save(temp_path)
            output_doc.close()
            
            # Print het bestand met verschillende methodes
            success = False
            error_msg = ""
            
            try:
                # Methode 1: Probeer via os.startfile (standaard Windows methode)
                for _ in range(copies):
                    os.startfile(temp_path, "print")
                success = True
                
            except Exception as e:
                error_msg = str(e)
                
                # Methode 2: Probeer via shell execute
                try:
                    import win32api
                    import win32print
                    
                    # Krijg standaard printer als geen printer gespecificeerd
                    if printer == "Standaard printer":
                        printer = win32print.GetDefaultPrinter()
                    
                    for _ in range(copies):
                        win32api.ShellExecute(
                            0,
                            "print",
                            temp_path,
                            f'/d:"{printer}"',
                            ".",
                            0
                        )
                    success = True
                    
                except Exception as e2:
                    error_msg = f"{error_msg}\n{str(e2)}"
                    
                    # Methode 3: Probeer via Adobe Reader of andere PDF reader
                    try:
                        # Zoek naar geïnstalleerde PDF readers
                        pdf_readers = []
                        
                        # Adobe Acrobat Reader
                        acrobat_paths = [
                            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                        ]
                        
                        # SumatraPDF
                        sumatra_paths = [
                            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
                        ]
                        
                        # Foxit Reader
                        foxit_paths = [
                            r"C:\Program Files\Foxit Software\Foxit Reader\FoxitReader.exe",
                            r"C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe",
                        ]
                        
                        reader_path = None
                        for path in acrobat_paths + sumatra_paths + foxit_paths:
                            if os.path.exists(path):
                                reader_path = path
                                break
                        
                        if reader_path:
                            # Print via gevonden PDF reader
                            for _ in range(copies):
                                subprocess.run([reader_path, "/t", temp_path, printer], 
                                             creationflags=subprocess.CREATE_NO_WINDOW)
                            success = True
                        else:
                            # Geen PDF reader gevonden, open het bestand gewoon
                            os.startfile(temp_path)
                            success = True
                            
                    except Exception as e3:
                        error_msg = f"{error_msg}\n{str(e3)}"
            
            if success:
                fit_text = " (passend gemaakt)" if fit_to_page else ""
                
                # Maak leesbare pagina lijst
                if len(pages) == 1:
                    page_info = f"Pagina {pages[0] + 1}"
                elif len(pages) <= 5:
                    page_info = f"Pagina's {', '.join(str(p + 1) for p in pages)}"
                else:
                    page_info = f"{len(pages)} pagina's"
                
                self.status_label.config(text=f"{len(pages)} pagina(s) naar printer gestuurd{fit_text}")
                messagebox.showinfo("Afdrukken", 
                    f"Document wordt afgedrukt:\n\n" +
                    f"Printer: {printer}\n" +
                    f"{page_info}\n" +
                    f"Kopieën: {copies}\n" +
                    f"Passend maken: {'Ja' if fit_to_page else 'Nee'}")
            else:
                # Toon duidelijke foutmelding met oplossing
                messagebox.showerror("Print Fout", 
                    "Kan niet automatisch printen.\n\n"
                    "Mogelijke oorzaken:\n"
                    "• Geen PDF reader geïnstalleerd of gekoppeld aan .pdf bestanden\n"
                    "• Geen printer geïnstalleerd\n\n"
                    "Oplossing:\n"
                    "1. Installeer Adobe Reader of een andere PDF reader\n"
                    "2. Stel deze in als standaard applicatie voor PDF bestanden\n"
                    "3. Of exporteer de pagina's en print deze handmatig\n\n"
                    f"Tijdelijk bestand opgeslagen op:\n{temp_path}")
            
            self.root.after(10000, lambda: self.cleanup_temp_file(temp_path))
                
        except Exception as e:
            messagebox.showerror("Print Fout", f"Kan niet printen:\n{str(e)}")

    def cleanup_temp_file(self, filepath):
        """Verwijder tijdelijk bestand"""
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass

    def save_form_data(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Bestanden", "*.pdf"), ("Alle Bestanden", "*.*")]
            )
            if save_path:
                try:
                    # Sla alle form data op in het PDF document
                    for page_num in range(len(tab.pdf_document)):
                        page = tab.pdf_document[page_num]
                        for widget in page.widgets():
                            if widget.field_name in tab.form_data:
                                widget.field_value = tab.form_data[widget.field_name]
                                widget.update()
                    
                    tab.pdf_document.save(save_path)
                    messagebox.showinfo("Succes", "PDF met ingevulde formuliervelden opgeslagen")
                except Exception as e:
                    messagebox.showerror("Fout", f"Kan PDF niet opslaan:\n{str(e)}")
    
    def split_pdf(self):
        """Splits PDF in losse pagina's"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Vraag output folder
        folder_path = filedialog.askdirectory(
            title="Selecteer map voor gesplitste pagina's"
        )
        
        if not folder_path:
            return
        
        try:
            base_name = os.path.splitext(os.path.basename(tab.file_path))[0]
            
            # Splits elke pagina
            for page_num in range(len(tab.pdf_document)):
                output_doc = fitz.open()
                output_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                
                output_path = os.path.join(folder_path, f"{base_name}_pagina_{page_num + 1}.pdf")
                output_doc.save(output_path)
                output_doc.close()
            
            messagebox.showinfo("Succes",
                f"PDF succesvol gesplitst!\n\n"
                f"Aantal pagina's: {len(tab.pdf_document)}\n"
                f"Opgeslagen in: {folder_path}")
            
            # Open de folder
            if messagebox.askyesno("Map Openen", "Wil je de map met bestanden openen?"):
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", folder_path])
                else:
                    subprocess.run(["xdg-open", folder_path])
            
        except Exception as e:
            messagebox.showerror("Fout", f"Kan PDF niet splitsen:\n{str(e)}")
    def show_pdf_info(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            metadata = tab.pdf_document.metadata
            info_text = (
                f"Titel: {metadata.get('title', 'N/A')}\n"
                f"Auteur: {metadata.get('author', 'N/A')}\n"
                f"Onderwerp: {metadata.get('subject', 'N/A')}\n"
                f"Trefwoorden: {metadata.get('keywords', 'N/A')}\n"
                f"Creator: {metadata.get('creator', 'N/A')}\n"
                f"Producer: {metadata.get('producer', 'N/A')}\n"
                f"Gemaakt: {metadata.get('creationDate', 'N/A')}\n"
                f"Gewijzigd: {metadata.get('modDate', 'N/A')}\n"
                f"Pagina's: {len(tab.pdf_document)}\n"
                f"Bestandsgrootte: {os.path.getsize(tab.file_path) / 1024:.1f} KB"
            )
            messagebox.showinfo("PDF Informatie", info_text)

    def show_edit_menu(self):
        """Toon bewerkingsmenu met opties"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak popup menu
        edit_menu = tk.Toplevel(self.root)
        edit_menu.title("PDF Bewerken")
        edit_menu.geometry("400x450")
        edit_menu.configure(bg=self.theme["BG_PRIMARY"])
        edit_menu.transient(self.root)
        edit_menu.grab_set()
        edit_menu.resizable(False, False)
        
        # Probeer icoon te laden
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                edit_menu.iconbitmap(icon_path)
        except:
            pass
        
        # Header
        header = tk.Frame(edit_menu, bg=self.theme["ACCENT_COLOR"], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="📝 PDF Bewerken", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content
        content = tk.Frame(edit_menu, bg=self.theme["BG_PRIMARY"])
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Opties
        tk.Label(content, text="Kies een bewerkingsoptie:", font=("Segoe UI", 10, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(0, 15))
        
        # Pagina's exporteren
        export_frame = self.create_menu_option(content, 
                                               "📄 Pagina's Exporteren",
                                               "Sla geselecteerde pagina's op als nieuw PDF bestand",
                                               lambda: [edit_menu.destroy(), self.export_pages()])
        export_frame.pack(fill=tk.X, pady=5)
        
        # PDF's samenvoegen
        merge_frame = self.create_menu_option(content,
                                              "📑 PDF's Samenvoegen",
                                              "Voeg meerdere PDF bestanden samen tot één",
                                              lambda: [edit_menu.destroy(), self.merge_pdfs()])
        merge_frame.pack(fill=tk.X, pady=5)
        
        # Pagina roteren
        rotate_frame = self.create_menu_option(content,
                                              "🔄 Pagina Roteren",
                                              "Roteer geselecteerde pagina's 90° / 180° / 270°",
                                              lambda: [edit_menu.destroy(), self.rotate_pages()])
        rotate_frame.pack(fill=tk.X, pady=5)
        
        # Footer met knop (moderne stijl)
        footer_frame = tk.Frame(edit_menu, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        tk.Button(footer_frame, text="Sluiten", command=edit_menu.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(pady=15)

    def create_menu_option(self, parent, title, description, command):
        """Maak een menu optie frame"""
        frame = tk.Frame(parent, bg=self.theme["BG_SECONDARY"], 
                        highlightbackground=self.theme["TEXT_SECONDARY"],
                        highlightthickness=1, cursor="hand2")
        
        inner = tk.Frame(frame, bg=self.theme["BG_SECONDARY"])
        inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        tk.Label(inner, text=title, font=("Segoe UI", 10, "bold"),
                bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                anchor="w").pack(fill=tk.X)
        
        tk.Label(inner, text=description, font=("Segoe UI", 8),
                bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_SECONDARY"],
                anchor="w", wraplength=340).pack(fill=tk.X, pady=(5, 0))
        
        frame.bind("<Button-1>", lambda e: command())
        for widget in frame.winfo_children():
            widget.bind("<Button-1>", lambda e: command())
            for child in widget.winfo_children():
                child.bind("<Button-1>", lambda e: command())
        
        def on_enter(e):
            frame.config(bg=self.theme["ACCENT_COLOR"])
            inner.config(bg=self.theme["ACCENT_COLOR"])
            for w in inner.winfo_children():
                w.config(bg=self.theme["ACCENT_COLOR"], fg="white")
        
        def on_leave(e):
            frame.config(bg=self.theme["BG_SECONDARY"])
            inner.config(bg=self.theme["BG_SECONDARY"])
            for i, w in enumerate(inner.winfo_children()):
                w.config(bg=self.theme["BG_SECONDARY"])
                if i == 0:
                    w.config(fg=self.theme["TEXT_PRIMARY"])
                else:
                    w.config(fg=self.theme["TEXT_SECONDARY"])
        
        frame.bind("<Enter>", on_enter)
        frame.bind("<Leave>", on_leave)
        
        return frame

    def export_pages(self):
        """Exporteer geselecteerde pagina's"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak moderne dialoog met header
        dialog = tk.Toplevel(self.root)
        dialog.title("Pagina's Exporteren")
        dialog.geometry("500x480")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📄 Pagina's Exporteren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Document info
        tk.Label(content_frame, 
                text=f"Document: {os.path.basename(tab.file_path)}\nTotaal aantal pagina's: {len(tab.pdf_document)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_SECONDARY"],
                justify=tk.LEFT).pack(pady=(0, 20), anchor="w")
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's wilt u exporteren?", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        pages_var = tk.StringVar()
        entry = tk.Entry(content_frame, textvariable=pages_var, font=("Segoe UI", 10), width=40)
        entry.pack(pady=8, fill=tk.X)
        entry.focus()
        
        # Voorbeelden in een nette box
        examples_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], 
                                 relief="flat", bd=1)
        examples_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(examples_frame, text="Voorbeelden:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", padx=15, pady=(10, 5))
        
        examples = [
            "• 1,3,5 (specifieke pagina's)",
            "• 1-5 (bereik)",
            "• 1-3,5,7-9 (combinatie)"
        ]
        
        for example in examples:
            tk.Label(examples_frame, text=example, 
                    font=("Segoe UI", 9),
                    bg=self.theme["BG_SECONDARY"], 
                    fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w", padx=25, pady=2)
        
        tk.Label(examples_frame, text=" ", bg=self.theme["BG_SECONDARY"]).pack(pady=5)
        
        def do_export():
            pages_input = pages_var.get()
            
            if not pages_input:
                messagebox.showwarning("Geen invoer", "Voer pagina nummers in")
                return
            
            # Parse pagina's
            pages = self.parse_page_range(pages_input, len(tab.pdf_document))
            
            if not pages:
                messagebox.showerror("Ongeldige invoer", "Ongeldige pagina selectie!")
                return
            
            # Vraag opslag locatie
            save_path = filedialog.asksaveasfilename(
                title="Exporteer pagina's als",
                defaultextension=".pdf",
                filetypes=[("PDF Bestanden", "*.pdf"), ("Alle Bestanden", "*.*")]
            )
            
            if not save_path:
                return
            
            try:
                # Maak nieuw PDF document
                new_doc = fitz.open()
                
                for page_num in pages:
                    new_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                
                new_doc.save(save_path)
                new_doc.close()
                
                dialog.destroy()
                messagebox.showinfo("Succes", 
                    f"{len(pages)} pagina('s) succesvol geëxporteerd naar:\n{os.path.basename(save_path)}")
                
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet exporteren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Exporteren", command=do_export,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Enter key binding
        entry.bind("<Return>", lambda e: do_export())

    def rotate_pages(self):
        """Roteer geselecteerde pagina's"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak rotatie dialoog
        rotate_dialog = tk.Toplevel(self.root)
        rotate_dialog.title("Pagina's Roteren")
        rotate_dialog.geometry("500x420")
        rotate_dialog.configure(bg=self.theme["BG_PRIMARY"])
        rotate_dialog.transient(self.root)
        rotate_dialog.grab_set()
        rotate_dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                rotate_dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(rotate_dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔄 Pagina's Roteren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(rotate_dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's?", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        page_var = tk.StringVar(value=f"{tab.current_page + 1}")
        tk.Entry(content_frame, textvariable=page_var, font=("Segoe UI", 10),
                width=40).pack(pady=8, fill=tk.X)
        
        tk.Label(content_frame, text="(bijv: 1,3,5 of 1-5)", font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w")
        
        # Rotatie hoek
        tk.Label(content_frame, text="Rotatie:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(15, 5))
        
        rotation_var = tk.IntVar(value=90)
        
        for angle in [90, 180, 270]:
            tk.Radiobutton(content_frame, text=f"{angle}° (rechtsom)" if angle == 90 else f"{angle}°",
                          variable=rotation_var, value=angle,
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=2)
        
        def do_rotate():
            pages_str = page_var.get()
            pages = self.parse_page_range(pages_str, len(tab.pdf_document))
            
            if not pages:
                messagebox.showerror("Ongeldige invoer", "Ongeldige pagina selectie!")
                return
            
            rotation = rotation_var.get()
            
            try:
                for page_num in pages:
                    page = tab.pdf_document[page_num]
                    page.set_rotation(rotation)
                
                # Ververs weergave
                self.display_page(tab)
                rotate_dialog.destroy()
                
                messagebox.showinfo("Succes", 
                    f"{len(pages)} pagina('s) geroteerd met {rotation}°\n\n" +
                    "Vergeet niet op te slaan om wijzigingen te behouden!")
                
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet roteren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(rotate_dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Roteren", command=do_rotate,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=rotate_dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)

    def exit_application(self):
        # Sla window geometry en state op voordat we afsluiten
        try:
            self.update_settings['window_geometry'] = self.root.geometry()
            self.update_settings['window_state'] = self.root.state()
            self.save_update_settings()
        except:
            pass  # Als opslaan mislukt, sluit gewoon af
        
        num_tabs = sum(1 for tab_id in self.notebook.tabs() 
                      if isinstance(self.notebook.nametowidget(tab_id), PDFTab))

        if num_tabs > 1:
            answer = messagebox.askyesno(
                "Afsluiten bevestigen",
                f"Er zijn {num_tabs} documenten geopend. Weet u zeker dat u wilt afsluiten?"
            )
            if not answer:
                return

        for tab_id in self.notebook.tabs():
            tab = self.notebook.nametowidget(tab_id)
            if isinstance(tab, PDFTab):
                tab.close_document()
        
        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def extract_pages(self):
        """Extraheer specifieke pagina's naar een nieuwe PDF"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showwarning("Geen document", "Open eerst een PDF document")
            return
        
        # Maak moderne dialoog met header
        dialog = tk.Toplevel(self.root)
        dialog.title("Pagina's extraheren")
        dialog.geometry("500x380")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📄 Pagina's Extraheren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Document info
        tk.Label(content_frame, text=f"Document: {os.path.basename(tab.file_path)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 20))
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's?", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        pages_var = tk.StringVar()
        tk.Entry(content_frame, textvariable=pages_var, font=("Segoe UI", 10),
                width=40).pack(pady=5, fill=tk.X)
        
        tk.Label(content_frame, text="(bijv: 1,3,5 of 1-5)", font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w")
        
        # Bereik selectie
        tk.Label(content_frame, text="Of selecteer bereik:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(15, 5))
        
        from_var = tk.StringVar(value="1")
        to_var = tk.StringVar(value=str(len(tab.pdf_document)))
        
        # Bereik layout
        range_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        range_frame.pack(anchor="w", pady=5)
        
        tk.Label(range_frame, text="Van:", font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=from_var, width=8,
                font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 15))
        
        tk.Label(range_frame, text="Tot:", font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=to_var, width=8,
                font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5)
        
        def do_extract():
            try:
                # Bepaal welke pagina's te extraheren
                if pages_var.get():
                    pages = self.parse_page_range(pages_var.get(), len(tab.pdf_document))
                else:
                    from_page = int(from_var.get()) - 1
                    to_page = int(to_var.get()) - 1
                    pages = list(range(from_page, to_page + 1))
                
                if not pages:
                    messagebox.showerror("Fout", "Ongeldige pagina selectie")
                    return
                
                # Vraag waar op te slaan
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Bestanden", "*.pdf")]
                )
                
                if save_path:
                    # Maak nieuwe PDF met geselecteerde pagina's
                    new_doc = fitz.open()
                    for page_num in pages:
                        new_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                    
                    new_doc.save(save_path)
                    new_doc.close()
                    
                    dialog.destroy()
                    messagebox.showinfo("Succes", 
                        f"{len(pages)} pagina's geëxtraheerd naar:\n{os.path.basename(save_path)}")
                    
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet extraheren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Extraheren", command=do_extract,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)

    def merge_pdfs(self):
        """Combineer meerdere PDF bestanden"""
        dialog = tk.Toplevel(self.root)
        dialog.title("PDF's combineren")
        dialog.geometry("550x570")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📑 PDF's Samenvoegen", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Optie voor openstaande bestanden
        option_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        def add_open_tabs():
            """Voeg alle openstaande PDF's toe aan de lijst"""
            added_count = 0
            for tab_id in self.notebook.tabs():
                tab = self.notebook.nametowidget(tab_id)
                if isinstance(tab, PDFTab):
                    if tab.file_path not in pdf_files:
                        pdf_files.append(tab.file_path)
                        listbox.insert(tk.END, os.path.basename(tab.file_path))
                        added_count += 1
            
            if added_count > 0:
                messagebox.showinfo("Toegevoegd", 
                    f"{added_count} openstaande PDF{'s' if added_count > 1 else ''} toegevoegd aan de lijst")
            else:
                messagebox.showinfo("Info", "Alle openstaande PDF's staan al in de lijst")
        
        tk.Button(option_frame, text="📂 Voeg openstaande PDF's toe", command=add_open_tabs,
                 bg=self.theme["SUCCESS_COLOR"], fg="white",
                 font=("Segoe UI", 9, "bold"), padx=15, pady=8,
                 relief="flat", cursor="hand2").pack(anchor="w")
        
        # Lijst van bestanden
        tk.Label(content_frame, text="Geselecteerde PDF bestanden:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        listbox = tk.Listbox(content_frame, height=10, font=("Segoe UI", 9))
        listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        pdf_files = []
        
        def add_files():
            files = filedialog.askopenfilenames(
                title="Selecteer PDF bestanden",
                filetypes=[("PDF Bestanden", "*.pdf")]
            )
            for file in files:
                if file not in pdf_files:
                    pdf_files.append(file)
                    listbox.insert(tk.END, os.path.basename(file))
        
        def remove_file():
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                listbox.delete(index)
                pdf_files.pop(index)
        
        def move_up():
            selection = listbox.curselection()
            if selection and selection[0] > 0:
                index = selection[0]
                # Swap in list
                pdf_files[index], pdf_files[index-1] = pdf_files[index-1], pdf_files[index]
                # Swap in listbox
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index-1, item)
                listbox.selection_set(index-1)
        
        def move_down():
            selection = listbox.curselection()
            if selection and selection[0] < len(pdf_files) - 1:
                index = selection[0]
                # Swap in list
                pdf_files[index], pdf_files[index+1] = pdf_files[index+1], pdf_files[index]
                # Swap in listbox
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index+1, item)
                listbox.selection_set(index+1)
        
        # Knoppen voor lijst beheer
        list_btn_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        list_btn_frame.pack(pady=10)
        
        tk.Button(list_btn_frame, text="➕ Toevoegen", command=add_files,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="➖ Verwijderen", command=remove_file,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="⬆ Omhoog", command=move_up,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="⬇ Omlaag", command=move_down,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        
        def do_merge():
            if len(pdf_files) < 2:
                messagebox.showwarning("Niet genoeg bestanden", 
                    "Selecteer tenminste 2 PDF bestanden om te combineren")
                return
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Bestanden", "*.pdf")],
                title="Gecombineerde PDF opslaan als"
            )
            
            if save_path:
                try:
                    merged_doc = fitz.open()
                    for pdf_file in pdf_files:
                        pdf_doc = fitz.open(pdf_file)
                        merged_doc.insert_pdf(pdf_doc)
                        pdf_doc.close()
                    
                    merged_doc.save(save_path)
                    merged_doc.close()
                    
                    dialog.destroy()
                    messagebox.showinfo("Succes", 
                        f"{len(pdf_files)} PDF's gecombineerd naar:\n{os.path.basename(save_path)}")
                    
                    # Vraag of gebruiker het gecombineerde bestand wil openen
                    if messagebox.askyesno("Openen?", "Wilt u het gecombineerde bestand openen?"):
                        self.add_new_tab(save_path)
                        
                except Exception as e:
                    messagebox.showerror("Fout", f"Kan PDF's niet combineren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Combineren", command=do_merge,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)


    def show_about(self):
        """Toon Over dialoog"""
        about = tk.Toplevel(self.root)
        about.title("Over NVict Reader")
        about.geometry("450x600")
        about.configure(bg=self.theme["BG_PRIMARY"])
        about.transient(self.root)
        about.resizable(False, False)
        
        # Voeg favicon toe aan taakbalk
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                about.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(about, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="Over NVict Reader", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(about, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Logo in content (met witte achtergrond voor betere zichtbaarheid)
        try:
            logo_path = get_resource_path('Logo.png')
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                logo_image.thumbnail((80, 80), Image.Resampling.LANCZOS)
                
                # Maak witte achtergrond voor logo
                bg_size = 100
                background = Image.new('RGB', (bg_size, bg_size), 'white')
                
                # Centreer logo op witte achtergrond
                offset = ((bg_size - logo_image.size[0]) // 2, (bg_size - logo_image.size[1]) // 2)
                if logo_image.mode == 'RGBA':
                    background.paste(logo_image, offset, logo_image)
                else:
                    background.paste(logo_image, offset)
                
                logo_photo = ImageTk.PhotoImage(background)
                logo_label = tk.Label(content_frame, image=logo_photo, bg=self.theme["BG_PRIMARY"])
                logo_label.image = logo_photo  # Keep reference
                logo_label.pack(pady=(10, 15))
        except Exception as e:
            print(f"Logo laden mislukt: {e}")
            pass
        
        # Titel
        tk.Label(content_frame, text="NVict Reader", 
                font=("Segoe UI", 16, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(pady=(0, 5))
        
        tk.Label(content_frame, text=f"Versie {APP_VERSION}", 
                font=("Segoe UI", 10),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 5))
        
        tk.Label(content_frame, text=f"© {self.get_current_year()} NVict Service", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 20))
        
        # Features in een mooie box
        features_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], 
                                 relief="flat", bd=1)
        features_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(features_frame, text="Functies:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", padx=20, pady=(15, 5))
        
        feature_list = [
            "✓ PDF's openen en bekijken",
            "✓ Tekst selecteren en kopiëren",
            "✓ Formulieren invullen",
            "✓ Pagina's exporteren",
            "✓ PDF's samenvoegen",
            "✓ Pagina's roteren",
            "✓ Afdrukken met opties",
            "✓ Zoeken in documenten"
        ]
        
        for feature in feature_list:
            tk.Label(features_frame, text=feature, font=("Segoe UI", 9),
                    bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                    anchor="w").pack(anchor="w", padx=30, pady=2)
        
        tk.Label(features_frame, text=" ", bg=self.theme["BG_SECONDARY"]).pack(pady=5)
        
        def open_website():
            webbrowser.open("https://www.nvict.nl/software.html")
        
        link_label = tk.Label(content_frame, text="www.nvict.nl", 
                             font=("Segoe UI", 9, "underline"),
                             bg=self.theme["BG_PRIMARY"], 
                             fg=self.theme["ACCENT_COLOR"],
                             cursor="hand2")
        link_label.pack(pady=15)
        link_label.bind("<Button-1>", lambda e: open_website())
        
        # Footer met knop (moderne stijl)
        footer_frame = tk.Frame(about, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        tk.Button(footer_frame, text="Sluiten", command=about.destroy,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(pady=15)

    def set_as_default_pdf(self):
        """Prompt user to set NVict Reader as default PDF viewer"""
        DefaultPDFHandler.prompt_set_as_default(self.root)

    def check_for_updates(self, silent=False):
        """Controleer of er updates beschikbaar zijn"""
        try:
            # Download versie info van server
            with urllib.request.urlopen(UPDATE_CHECK_URL, timeout=5) as response:
                data = json.loads(response.read().decode('utf-8'))
                
                latest_version = data.get("version", "0.0")
                download_url = data.get("download_url", "")
                release_notes = data.get("release_notes", "")
                
                # Vergelijk versies
                current_parts = [int(x) for x in APP_VERSION.split('.')]
                latest_parts = [int(x) for x in latest_version.split('.')]
                
                # Pad version parts als ze verschillende lengtes hebben
                max_length = max(len(current_parts), len(latest_parts))
                current_parts += [0] * (max_length - len(current_parts))
                latest_parts += [0] * (max_length - len(latest_parts))
                
                update_available = latest_parts > current_parts
                
                if update_available:
                    self.show_update_dialog(latest_version, download_url, release_notes)
                else:
                    if not silent:
                        messagebox.showinfo("Geen updates", 
                            f"U gebruikt al de nieuwste versie ({APP_VERSION})")
                        
        except urllib.error.URLError:
            if not silent:
                messagebox.showerror("Verbindingsfout", 
                    "Kan niet verbinden met de update server.\n\n"
                    "Controleer uw internetverbinding en probeer het later opnieuw.")
        except Exception as e:
            if not silent:
                messagebox.showerror("Fout", 
                    f"Fout bij controleren op updates:\n{str(e)}")

    def show_update_dialog(self, new_version, download_url, release_notes):
        """Toon dialoog met update informatie en automatische download/installatie optie"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Beschikbaar")
        dialog.geometry("520x550")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur
        header_frame = tk.Frame(dialog, bg=self.theme["SUCCESS_COLOR"], height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🎉 Update Beschikbaar!", 
                font=("Segoe UI", 16, "bold"),
                bg=self.theme["SUCCESS_COLOR"], fg="white").pack(pady=25)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Versie informatie
        version_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], relief="flat")
        version_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = f"Huidige versie: {APP_VERSION}\nNieuwe versie: {new_version}"
        tk.Label(version_frame, text=info_text, 
                font=("Segoe UI", 10),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"],
                justify=tk.LEFT).pack(padx=20, pady=15)
        
        # Release notes
        if release_notes:
            tk.Label(content_frame, text="Wat is er nieuw:", 
                    font=("Segoe UI", 10, "bold"),
                    bg=self.theme["BG_PRIMARY"], 
                    fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(0, 5))
            
            notes_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"])
            notes_frame.pack(fill=tk.BOTH, expand=True)
            
            notes_text = tk.Text(notes_frame, wrap=tk.WORD, 
                               font=("Segoe UI", 9),
                               bg=self.theme["BG_SECONDARY"],
                               fg=self.theme["TEXT_PRIMARY"],
                               relief="flat", height=8)
            notes_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            notes_text.insert("1.0", release_notes)
            notes_text.config(state=tk.DISABLED)
        
        # Installatie instructie
        install_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        install_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(install_frame, 
                text="ℹ️ Na download wordt de installer automatisch geopend.\nSluit eerst NVict Reader af voordat u installeert.",
                font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_SECONDARY"],
                justify=tk.LEFT).pack(anchor="w")
        
        # Footer met knoppen
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=80)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        def download_and_install():
            """Download update en start installatie automatisch"""
            if download_url:
                dialog.destroy()
                self.download_and_install_update(download_url, new_version)
        
        def download_only():
            """Open alleen de download pagina in browser"""
            if download_url:
                webbrowser.open(download_url)
                dialog.destroy()
        
        tk.Button(btn_container, text="Download & Installeer", command=download_and_install,
                 bg=self.theme["SUCCESS_COLOR"], fg="white",
                 font=("Segoe UI", 10, "bold"), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
        
        tk.Button(btn_container, text="Alleen Download", command=download_only,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
        
        tk.Button(btn_container, text="Later", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
    
    def download_and_install_update(self, download_url, version):
        """Download update en start installatie automatisch"""
        try:
            # Toon voortgang dialoog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Update Downloaden")
            progress_dialog.geometry("400x150")
            progress_dialog.configure(bg=self.theme["BG_PRIMARY"])
            progress_dialog.transient(self.root)
            progress_dialog.resizable(False, False)
            
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    progress_dialog.iconbitmap(icon_path)
            except:
                pass
            
            tk.Label(progress_dialog, text="Update downloaden...", 
                    font=("Segoe UI", 12, "bold"),
                    bg=self.theme["BG_PRIMARY"], 
                    fg=self.theme["TEXT_PRIMARY"]).pack(pady=20)
            
            status_label = tk.Label(progress_dialog, text="Bezig met downloaden...",
                                   font=("Segoe UI", 9),
                                   bg=self.theme["BG_PRIMARY"],
                                   fg=self.theme["TEXT_SECONDARY"])
            status_label.pack(pady=10)
            
            progress_dialog.update()
            
            # Download in achtergrond thread
            def download_thread():
                try:
                    # Download naar temp directory
                    temp_dir = tempfile.gettempdir()
                    filename = f"NVict_Reader_v{version}_Setup.exe"
                    filepath = os.path.join(temp_dir, filename)
                    
                    # Download bestand
                    urllib.request.urlretrieve(download_url, filepath)
                    
                    # Update UI in main thread
                    self.root.after(0, lambda: self._finish_download(progress_dialog, filepath))
                    
                except Exception as e:
                    self.root.after(0, lambda: self._download_error(progress_dialog, str(e)))
            
            thread = threading.Thread(target=download_thread, daemon=True)
            thread.start()
            
        except Exception as e:
            messagebox.showerror("Download Fout", 
                f"Kan update niet downloaden:\n{str(e)}\n\nProbeer handmatig te downloaden via de website.")
    
    def _finish_download(self, progress_dialog, filepath):
        """Voltooi download en start installer"""
        try:
            progress_dialog.destroy()
            
            if os.path.exists(filepath):
                # Vraag bevestiging om installer te starten
                if messagebox.askyesno("Update Downloaden Voltooid",
                    f"Update succesvol gedownload!\n\n"
                    f"Wilt u de installer nu starten?\n\n"
                    f"Let op: Sluit eerst NVict Reader af voordat u de installatie voltooit."):
                    
                    # Start installer
                    if platform.system() == "Windows":
                        os.startfile(filepath)
                    elif platform.system() == "Darwin":
                        subprocess.run(["open", filepath])
                    else:
                        subprocess.run(["xdg-open", filepath])
                    
                    # Sluit applicatie
                    self.root.after(1000, self.exit_application)
            else:
                messagebox.showerror("Fout", "Download bestand niet gevonden")
                
        except Exception as e:
            messagebox.showerror("Fout", f"Kan installer niet starten:\n{str(e)}")
    
    def _download_error(self, progress_dialog, error_msg):
        """Toon download fout"""
        progress_dialog.destroy()
        messagebox.showerror("Download Fout",
            f"Kan update niet downloaden:\n{error_msg}\n\n"
            f"Probeer handmatig te downloaden via de website.")

    def run(self):
        self.notebook.add(self.welcome_frame)
        # WM_DELETE_WINDOW wordt nu in main() ingesteld met single instance cleanup
        
        # Start automatische update check in achtergrond (na 2 seconden)
        self.root.after(2000, self.check_for_updates_on_startup)
        
        self.root.mainloop()

def main():
    try:
        # Controleer voor single instance
        single_instance = SingleInstance()
        
        # Check of er een bestand is meegegeven als argument
        file_to_open = None
        print_mode = False
        
        # Parse command line argumenten
        if len(sys.argv) > 1:
            if sys.argv[1] == "--print" and len(sys.argv) > 2:
                # Format: NVictReader.exe --print "bestand.pdf"
                print_mode = True
                if os.path.exists(sys.argv[2]):
                    file_to_open = os.path.abspath(sys.argv[2])
            elif os.path.exists(sys.argv[1]):
                # Format: NVictReader.exe "bestand.pdf"
                file_to_open = os.path.abspath(sys.argv[1])
        
        # Check of er al een instance draait
        if single_instance.is_already_running():
            # Stuur bestand naar bestaande instance als er een is
            if file_to_open:
                if single_instance.send_to_existing_instance(file_to_open):
                    print(f"Bestand verzonden naar bestaande instance: {file_to_open}")
                    return  # Sluit deze instance af
                else:
                    print("Kon niet communiceren met bestaande instance, start nieuwe instance")
            else:
                # Geen bestand om te openen, gewoon een nieuwe instance starten
                # (gebruiker wil misschien een tweede venster)
                pass
        
        # Start de applicatie
        app = NVictReader()
        
        # Start single instance server
        single_instance.start_server(app)
        
        # Open bestand als er een is meegegeven
        if file_to_open:
            app.root.after(100, lambda: app.add_new_tab(file_to_open))
            
            # Als print mode, open automatisch het print dialoog
            if print_mode:
                app.root.after(500, lambda: app.print_pdf())
        
        # Zorg dat server wordt gestopt bij afsluiten
        def on_closing():
            single_instance.stop()
            app.exit_application()
        
        app.root.protocol("WM_DELETE_WINDOW", on_closing)
        
        app.run()
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("\nDruk op Enter om af te sluiten...")

if __name__ == "__main__":
    main()

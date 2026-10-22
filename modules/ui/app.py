"""
Modern shell with sidebar navigation - Phase 3

Usage:
    from modules.ui.app import ModernApp
    
    app = ModernApp()
    app.mainloop()
"""
import sys
import os
from pathlib import Path

# Add project root to path when run directly
if __name__ == "__main__":
    project_root = Path(__file__).parent.parent.parent
    sys.path.insert(0, str(project_root))

import customtkinter as ctk
from modules.ui.theme import apply_theme, color, font, SIZES, SPACING, RADIUS
from modules.ui.dispatcher import UiDispatcher, EventEmitter
from modules.ui.components.toast import ToastQueue

class ModernApp(ctk.CTk):
    """
    Main application shell with sidebar navigation.
    
    Replaces the old CTkTabview approach.
    """
    
    def __init__(self):
        super().__init__()
        
        # Apply theme before creating widgets
        apply_theme("dark")
        
        # Window setup
        self.title("Excel Data Processor - ODOO Test")
        self.geometry("1200x800")
        self.minsize(960, 640)
        
        # Configure root background
        self.configure(fg_color=color("bg"))
        
        # Thread-safe dispatcher
        self.dispatcher = UiDispatcher(self)
        self.emitter = EventEmitter(self.dispatcher)
        
        # Toast notifications
        self.toast = ToastQueue(self)
        
        # Current page tracking
        self.current_page = None
        self.pages = {}
        
        # Build UI
        self._build_shell()
        
        # Keyboard shortcuts
        self._bind_shortcuts()
        
        # Close handler
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _build_shell(self):
        """Build the main shell layout"""
        # Main container (3 sections: sidebar, content, status)
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True)
        
        # Sidebar
        self.sidebar = self._build_sidebar(main_container)
        self.sidebar.pack(side="left", fill="y")
        
        # Content area (right side)
        right_panel = ctk.CTkFrame(main_container, fg_color="transparent")
        right_panel.pack(side="left", fill="both", expand=True)
        
        # Page header
        self.header = self._build_header(right_panel)
        self.header.pack(fill="x", padx=SPACING["xl"], pady=(SPACING["xl"], SPACING["md"]))
        
        # Page container
        self.page_container = ctk.CTkFrame(right_panel, fg_color="transparent")
        self.page_container.pack(fill="both", expand=True, 
                                padx=SPACING["xl"], pady=(0, SPACING["md"]))
        
        # Status bar
        self.status_bar = self._build_status_bar(right_panel)
        self.status_bar.pack(fill="x", padx=SPACING["xl"], pady=(0, SPACING["md"]))
    
    def _build_sidebar(self, parent) -> ctk.CTkFrame:
        """Build sidebar with navigation"""
        sidebar = ctk.CTkFrame(
            parent,
            width=SIZES["sidebar_width"],
            fg_color=color("surface"),
            corner_radius=0
        )
        sidebar.pack_propagate(False)
        
        # Logo/Title
        title_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        title_frame.pack(fill="x", padx=SPACING["lg"], pady=SPACING["xl"])
        
        logo_label = ctk.CTkLabel(
            title_frame,
            text="▣",
            font=("Arial", 24),
            text_color=color("primary")
        )
        logo_label.pack(side="left")
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Excel\nProcessor",
            font=("Inter", 14, "bold"),
            justify="left"
        )
        title_label.pack(side="left", padx=(SPACING["sm"], 0))
        
        # Navigation buttons
        nav_container = ctk.CTkFrame(sidebar, fg_color="transparent")
        nav_container.pack(fill="both", expand=True, padx=SPACING["md"])
        
        self.nav_buttons = {}
        nav_items = [
            ("dashboard", "Dashboard", "📊"),
            ("files", "Files", "📁"),
            ("database", "Database", "🗄"),
            ("rencana", "Rencana", "📝"),
            ("logs", "Logs", "📋"),
        ]
        
        for page_id, label, icon in nav_items:
            btn = self._create_nav_button(nav_container, page_id, label, icon)
            btn.pack(fill="x", pady=(0, SPACING["xs"]))
            self.nav_buttons[page_id] = btn
        
        # Bottom section (sync status + theme toggle)
        bottom = ctk.CTkFrame(sidebar, fg_color="transparent")
        bottom.pack(fill="x", side="bottom", padx=SPACING["lg"], pady=SPACING["lg"])
        
        # Sync status
        from modules.ui.components import StatusBadge
        self.sync_badge = StatusBadge(bottom, text="Sync: Off", state="idle")
        self.sync_badge.pack(fill="x", pady=(0, SPACING["md"]))
        
        # Theme toggle
        self.theme_toggle = ctk.CTkSwitch(
            bottom,
            text="🌙 Dark",
            command=self._toggle_theme
        )
        self.theme_toggle.pack(fill="x")
        
        return sidebar
    
    def _create_nav_button(self, parent, page_id: str, label: str, icon: str) -> ctk.CTkButton:
        """Create navigation button"""
        btn = ctk.CTkButton(
            parent,
            text=f"{icon}  {label}",
            anchor="w",
            height=40,
            corner_radius=RADIUS["input"],
            fg_color="transparent",
            text_color=color("text"),
            hover_color=color("surface_alt"),
            command=lambda: self.show_page(page_id)
        )
        return btn
    
    def _build_header(self, parent) -> ctk.CTkFrame:
        """Build page header"""
        header = ctk.CTkFrame(parent, fg_color="transparent")
        
        self.page_title = ctk.CTkLabel(
            header,
            text="Dashboard",
            font=font("page_title"),
            anchor="w"
        )
        self.page_title.pack(side="left", fill="x", expand=True)
        
        return header
    
    def _build_status_bar(self, parent) -> ctk.CTkFrame:
        """Build status bar"""
        status = ctk.CTkFrame(
            parent,
            height=32,
            fg_color=color("surface"),
            corner_radius=RADIUS["input"]
        )
        
        self.status_text = ctk.CTkLabel(
            status,
            text="Ready",
            font=font("body"),
            text_color=color("muted")
        )
        self.status_text.pack(side="left", padx=SPACING["md"])
        
        # Progress bar (hidden by default)
        self.status_progress = ctk.CTkProgressBar(
            status,
            mode="indeterminate",
            height=4
        )
        # Don't pack initially
        
        return status
    
    def register_page(self, page_id: str, page_widget):
        """Register a page widget"""
        self.pages[page_id] = page_widget
        page_widget.grid(row=0, column=0, sticky="nsew", in_=self.page_container)
        self.page_container.grid_rowconfigure(0, weight=1)
        self.page_container.grid_columnconfigure(0, weight=1)
        page_widget.grid_remove()  # Hide initially
    
    def show_page(self, page_id: str):
        """Switch to a page"""
        # Hide current page
        if self.current_page and self.current_page in self.pages:
            self.pages[self.current_page].grid_remove()
        
        # Show new page
        if page_id in self.pages:
            self.pages[page_id].grid()
            self.current_page = page_id
            
            # Update nav button states
            for btn_id, btn in self.nav_buttons.items():
                if btn_id == page_id:
                    btn.configure(fg_color=color("primary"),
                                text_color=("#FFFFFF", "#FFFFFF"))
                else:
                    btn.configure(fg_color="transparent",
                                text_color=color("text"))
            
            # Update title
            titles = {
                "dashboard": "Dashboard",
                "files": "File Configuration",
                "database": "Database Configuration",
                "rencana": "Input Rencana Benda Uji",
                "logs": "Process Logs"
            }
            self.page_title.configure(text=titles.get(page_id, page_id.title()))
    
    def set_busy(self, busy: bool, text: str = "Processing..."):
        """Show/hide busy state"""
        if busy:
            self.status_text.configure(text=text)
            self.status_progress.pack(side="right", padx=SPACING["md"], fill="x", expand=True)
            self.status_progress.start()
        else:
            self.status_progress.stop()
            self.status_progress.pack_forget()
            self.status_text.configure(text="Ready")
    
    def _toggle_theme(self):
        """Toggle light/dark theme"""
        from modules.ui.theme import get_current_mode
        
        current = get_current_mode()
        new_mode = "light" if current == "dark" else "dark"
        apply_theme(new_mode)
        
        # Update toggle text
        icon = "☀" if new_mode == "light" else "🌙"
        mode_text = "Light" if new_mode == "light" else "Dark"
        self.theme_toggle.configure(text=f"{icon} {mode_text}")
    
    def _on_closing(self):
        """Handle window close"""
        # ponytail: check for running processes, show ConfirmDialog if needed
        self.dispatcher.stop()
        self.destroy()
    
    def _bind_shortcuts(self):
        """Bind keyboard shortcuts"""
        # Page navigation: Ctrl+1-5
        self.bind("<Control-Key-1>", lambda e: self.show_page("dashboard"))
        self.bind("<Control-Key-2>", lambda e: self.show_page("files"))
        self.bind("<Control-Key-3>", lambda e: self.show_page("database"))
        self.bind("<Control-Key-4>", lambda e: self.show_page("rencana"))
        self.bind("<Control-Key-5>", lambda e: self.show_page("logs"))
        
        # Theme toggle: Ctrl+T
        self.bind("<Control-t>", lambda e: self._toggle_theme())
        self.bind("<Control-T>", lambda e: self._toggle_theme())


# Demo/test function
if __name__ == "__main__":
    import tkinter as tk
    from modules.ui.pages import (
        DashboardPage,
        DatabasePage,
        FilesPage,
        LogsPage,
        RencanaPage
    )
    
    app = ModernApp()
    
    # Create StringVars for file/db configuration
    file_vars = {
        "rencana": tk.StringVar(value="Rencana/RENCANA.xlsx"),
        "pengujian": tk.StringVar(value="Pengujian/PENGUJIAN.xlsx"),
        "output": tk.StringVar(value="Output"),
        "sheet_source": tk.StringVar(value="Report Excel"),
        "sheet_odoo": tk.StringVar(value="ODOO"),
        "sheet_pengujian": tk.StringVar(value="pengujian"),
        "sheet_bjdt": tk.StringVar(value="bjdt_id"),
    }
    
    db_vars = {
        "user": tk.StringVar(value="postgres"),
        "password": tk.StringVar(value=""),
        "host": tk.StringVar(value="localhost"),
        "port": tk.StringVar(value="5432"),
        "name": tk.StringVar(value="alatujidb"),
        "table": tk.StringVar(value="pengujian"),
    }
    
    # Create real pages
    dashboard = DashboardPage(app.page_container, app)
    files = FilesPage(app.page_container, app, file_vars)
    database = DatabasePage(app.page_container, app, db_vars)
    rencana = RencanaPage(app.page_container, app)
    logs = LogsPage(app.page_container, app)
    
    # Register pages
    app.register_page("dashboard", dashboard)
    app.register_page("files", files)
    app.register_page("database", database)
    app.register_page("rencana", rencana)
    app.register_page("logs", logs)
    
    # Show first page
    app.show_page("dashboard")
    
    # Test toast
    app.after(1000, lambda: app.toast.success("Modern UI initialized"))
    
    app.mainloop()

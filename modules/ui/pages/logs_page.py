"""Logs page with colored console"""
import customtkinter as ctk
from modules.ui.theme import SPACING, RADIUS
from modules.ui.components import SectionCard, LogConsole, SecondaryButton
from tkinter import filedialog
from datetime import datetime

class LogsPage(ctk.CTkFrame):
    """
    Logs page with console and toolbar.
    
    Replaces the old "Process Logs" tab.
    """
    
    def __init__(self, parent, app):
        """
        Args:
            parent: Parent widget
            app: ModernApp instance
        """
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        
        self._build_ui()
    
    def _build_ui(self):
        """Build page content"""
        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, SPACING["md"]))
        
        # Left side - level filter
        level_label = ctk.CTkLabel(toolbar, text="Level:")
        level_label.pack(side="left", padx=(0, SPACING["sm"]))
        
        self.level_filter = ctk.CTkSegmentedButton(
            toolbar,
            values=["All", "Info", "Warning", "Error"],
            command=self._on_level_change
        )
        self.level_filter.set("All")
        self.level_filter.pack(side="left", padx=(0, SPACING["lg"]))
        
        # Right side - actions
        action_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        action_frame.pack(side="right")
        
        SecondaryButton(
            action_frame,
            text="Save",
            command=self._save_logs,
            width=80
        ).pack(side="left", padx=(0, SPACING["xs"]))
        
        SecondaryButton(
            action_frame,
            text="Clear",
            command=self._clear_logs,
            width=80
        ).pack(side="left", padx=(0, SPACING["xs"]))
        
        # Auto-scroll toggle
        self.autoscroll_var = ctk.BooleanVar(value=True)
        self.autoscroll_toggle = ctk.CTkSwitch(
            action_frame,
            text="Auto-scroll",
            variable=self.autoscroll_var,
            command=self._toggle_autoscroll
        )
        self.autoscroll_toggle.pack(side="left")
        
        # Console
        self.console = LogConsole(self)
        self.console.pack(fill="both", expand=True)
        
        # Demo logs
        self._add_demo_logs()
    
    def _on_level_change(self, level: str):
        """Handle level filter change"""
        # ponytail: implement filtering by storing all logs and rebuilding visible text
        pass
    
    def _save_logs(self):
        """Save logs to file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=(("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")),
            initialfile=f"process_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )
        
        if filename:
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(self.console.get_all_text())
                self.app.toast.success(f"Logs saved to {filename}")
            except Exception as e:
                from modules.ui.components import ErrorDialog
                ErrorDialog(self, "Error", "Failed to save logs", detail=str(e))
    
    def _clear_logs(self):
        """Clear all logs"""
        from modules.ui.components import ConfirmDialog
        
        if ConfirmDialog(self, "Clear Logs", "Clear all log entries?").wait_result():
            self.console.clear()
            self.app.toast.info("Logs cleared")
    
    def _toggle_autoscroll(self):
        """Toggle auto-scroll"""
        self.console.set_autoscroll(self.autoscroll_var.get())
    
    def append_log(self, message: str, level: str = "INFO"):
        """Append log message (thread-safe via dispatcher)"""
        self.app.dispatcher.post(self.console.append, message, level)
    
    def _add_demo_logs(self):
        """Add demo logs for testing"""
        self.console.append("[2026-10-13 16:30:00] Application started", "INFO")
        self.console.append("[2026-10-13 16:30:01] Loading configuration...", "INFO")
        self.console.append("[2026-10-13 16:30:02] Connected to database: alatujidb", "SUCCESS")
        self.console.append("[2026-10-13 16:30:03] Excel file loaded: 1,234 rows", "INFO")
        self.console.append("[2026-10-13 16:30:05] Warning: Missing BJDT ID for 5 entries", "WARNING")
        self.console.append("[2026-10-13 16:30:06] Processed 100 rows...", "INFO")
        self.console.append("[2026-10-13 16:30:07] Processed 200 rows...", "INFO")
        self.console.append("[2026-10-13 16:30:08] Network timeout, retrying...", "WARNING")
        self.console.append("[2026-10-13 16:30:10] Upload complete: 250 records", "SUCCESS")

"""Files configuration page"""
import customtkinter as ctk
import tkinter as tk
from modules.ui.theme import SPACING
from modules.ui.components import SectionCard, FileField, LabeledEntry

class FilesPage(ctk.CTkScrollableFrame):
    """
    Files configuration page.
    
    Replaces the old "File Configuration" tab.
    """
    
    def __init__(self, parent, app, file_vars: dict):
        """
        Args:
            parent: Parent widget
            app: ModernApp instance
            file_vars: Dict of StringVars {rencana, pengujian, output, sheets}
        """
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        self.file_vars = file_vars
        
        self._build_ui()
    
    def _build_ui(self):
        """Build page content"""
        # Files card
        files_card = SectionCard(
            self,
            title="Excel Files",
            subtitle="Select input and output files"
        )
        files_card.pack(fill="x", pady=(0, SPACING["xl"]))
        
        # File fields
        FileField(
            files_card.body,
            label="Rencana Excel File",
            variable=self.file_vars["rencana"],
            mode="file",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        ).pack(fill="x", pady=(0, SPACING["md"]))
        
        FileField(
            files_card.body,
            label="Pengujian Excel File",
            variable=self.file_vars["pengujian"],
            mode="file",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        ).pack(fill="x", pady=(0, SPACING["md"]))
        
        FileField(
            files_card.body,
            label="Output Directory",
            variable=self.file_vars["output"],
            mode="dir"
        ).pack(fill="x")
        
        # Sheets card
        sheets_card = SectionCard(
            self,
            title="Sheet Names",
            subtitle="Excel worksheet names to process"
        )
        sheets_card.pack(fill="x")
        
        # Sheet name form (2 columns)
        form = ctk.CTkFrame(sheets_card.body, fg_color="transparent")
        form.pack(fill="both", expand=True)
        
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        
        # Row 0: Source | ODOO
        LabeledEntry(
            form,
            label="Source Sheet",
            variable=self.file_vars["sheet_source"],
            helper="Default: Report Excel",
            width=300
        ).grid(row=0, column=0, sticky="ew", padx=(0, SPACING["md"]), pady=(0, SPACING["md"]))
        
        LabeledEntry(
            form,
            label="ODOO Sheet",
            variable=self.file_vars["sheet_odoo"],
            helper="Default: ODOO",
            width=300
        ).grid(row=0, column=1, sticky="ew", pady=(0, SPACING["md"]))
        
        # Row 1: Pengujian | BJDT ID
        LabeledEntry(
            form,
            label="Pengujian Sheet",
            variable=self.file_vars["sheet_pengujian"],
            helper="Default: pengujian",
            width=300
        ).grid(row=1, column=0, sticky="ew", padx=(0, SPACING["md"]))
        
        LabeledEntry(
            form,
            label="BJDT ID Sheet",
            variable=self.file_vars["sheet_bjdt"],
            helper="Default: bjdt_id",
            width=300
        ).grid(row=1, column=1, sticky="ew")

"""Dashboard page with process pipeline"""
import customtkinter as ctk
import threading
from modules.ui.theme import SPACING, color, RADIUS
from modules.ui.components import (
    SectionCard,
    PrimaryButton,
    SecondaryButton,
    SuccessButton,
    StatusBadge
)
from modules.controllers.process_controller import ProcessController, ProcessSettings

class DashboardPage(ctk.CTkScrollableFrame):
    """
    Dashboard page with process pipeline and actions.
    
    Replaces the old "Process Control" tab.
    """
    
    def __init__(self, parent, app):
        """
        Args:
            parent: Parent widget
            app: ModernApp instance
        """
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        self.controller = None
        self.step_badges = {}
        
        self._build_ui()
    
    def _build_ui(self):
        """Build page content"""
        # Pipeline card
        pipeline_card = SectionCard(
            self,
            title="Process Pipeline",
            subtitle="Automated data processing workflow"
        )
        pipeline_card.pack(fill="x", pady=(0, SPACING["xl"]))
        
        # Stepper
        stepper = ctk.CTkFrame(pipeline_card.body, fg_color="transparent")
        stepper.pack(fill="x", pady=(0, SPACING["lg"]))
        
        steps = [
            ("copy", "Copy Data", "idle"),
            ("bjdt", "Get BJDT ID", "idle"),
            ("csv", "Generate CSV", "idle"),
            ("upload", "Upload DB", "idle")
        ]
        
        for i, (step_id, step_name, state) in enumerate(steps):
            step_frame = ctk.CTkFrame(stepper, fg_color="transparent")
            step_frame.pack(side="left", expand=True)
            
            # Step indicator
            badge = StatusBadge(step_frame, text=str(i+1), state=state)
            badge.pack()
            self.step_badges[step_id] = badge
            
            # Step label
            label = ctk.CTkLabel(step_frame, text=step_name, font=("Inter", 11))
            label.pack(pady=(SPACING["xs"], 0))
            
            # Arrow (except last)
            if i < len(steps) - 1:
                arrow = ctk.CTkLabel(stepper, text="→", font=("Arial", 20),
                                    text_color=color("muted"))
                arrow.pack(side="left", padx=SPACING["sm"])
        
        # Full process button
        self.full_process_btn = PrimaryButton(
            pipeline_card.body,
            text="▶ Start Full Process",
            command=self._run_full_process,
            width=200
        )
        self.full_process_btn.pack()
        
        # Individual actions card
        actions_card = SectionCard(
            self,
            title="Individual Actions",
            subtitle="Run process steps separately"
        )
        actions_card.pack(fill="x", pady=(0, SPACING["xl"]))
        
        # Actions grid (3 columns)
        grid = ctk.CTkFrame(actions_card.body, fg_color="transparent")
        grid.pack(fill="x")
        
        actions = [
            ("📋 Copy Data", self._run_copy_only),
            ("🔢 Get BJDT ID", self._run_bjdt_only),
            ("📄 Generate CSV", self._run_csv_only),
            ("📤 Upload CSV", self._run_upload_csv),
            ("🗄 Upload to DB", self._run_upload_only),
            ("⚙ Beban Config", self._open_beban_config),
        ]
        
        for i, (text, cmd) in enumerate(actions):
            row = i // 3
            col = i % 3
            
            SecondaryButton(
                grid,
                text=text,
                command=cmd,
                width=180,
                height=50
            ).grid(row=row, column=col, padx=SPACING["sm"], pady=SPACING["sm"], sticky="ew")
        
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(2, weight=1)
        
        # Tools card
        tools_card = SectionCard(
            self,
            title="Tools",
            subtitle="Additional utilities"
        )
        tools_card.pack(fill="x")
        
        tools_frame = ctk.CTkFrame(tools_card.body, fg_color="transparent")
        tools_frame.pack(fill="x")
        
        SecondaryButton(
            tools_frame,
            text="🔧 Open Grid Benda Uji",
            command=self._open_grid,
            width=200
        ).pack(side="left", padx=(0, SPACING["sm"]))
    
    def _get_settings(self) -> ProcessSettings:
        """Build settings from app state"""
        # Get file/db vars from app (passed during page creation)
        # For now, use defaults - will be wired in app.py
        import os
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
        
        return ProcessSettings(
            rencana_path=os.path.join(base_dir, 'Rencana', 'RENCANA.xlsx'),
            pengujian_path=os.path.join(base_dir, 'Pengujian', 'PENGUJIAN.xlsx'),
            output_dir=os.path.join(base_dir, 'Output'),
            db_host="localhost",
            db_port="5432",
            db_user="postgres",
            db_password="",
            db_name="alatujidb",
            db_table="pengujian"
        )
    
    def _reset_badges(self):
        """Reset all step badges to idle"""
        for badge in self.step_badges.values():
            badge.set_state("idle")
    
    def _on_step_started(self, step_name: str):
        """Callback when step starts"""
        step_map = {
            "Copy Data": "copy",
            "Get BJDT ID": "bjdt",
            "Generate CSV": "csv",
            "Upload DB": "upload"
        }
        
        step_id = step_map.get(step_name)
        if step_id and step_id in self.step_badges:
            self.step_badges[step_id].set_state("running", "...")
    
    def _on_step_completed(self, step_name: str, success: bool):
        """Callback when step completes"""
        step_map = {
            "Copy Data": "copy",
            "Get BJDT ID": "bjdt",
            "Generate CSV": "csv",
            "Upload DB": "upload"
        }
        
        step_id = step_map.get(step_name)
        if step_id and step_id in self.step_badges:
            state = "success" if success else "error"
            text = "✓" if success else "✗"
            self.step_badges[step_id].set_state(state, text)
    
    def _on_log(self, message: str, level: str):
        """Callback for log messages"""
        # Find logs page and append
        if "logs" in self.app.pages:
            logs_page = self.app.pages["logs"]
            logs_page.append_log(message, level)
    
    def _run_full_process(self):
        """Run full pipeline"""
        self.full_process_btn.configure(state="disabled", text="Processing...")
        self._reset_badges()
        self.app.set_busy(True, "Running full process...")
        
        def worker():
            settings = self._get_settings()
            controller = ProcessController(
                settings,
                on_step_started=lambda step: self.app.dispatcher.post(self._on_step_started, step),
                on_step_completed=lambda step, ok: self.app.dispatcher.post(self._on_step_completed, step, ok),
                on_log=lambda msg, lvl: self.app.dispatcher.post(self._on_log, msg, lvl)
            )
            
            result = controller.run_full_process()
            
            def finish():
                self.app.set_busy(False)
                self.full_process_btn.configure(state="normal", text="▶ Start Full Process")
                
                if result.success:
                    self.app.toast.success(f"Process complete! {result.steps_completed}/4 steps")
                else:
                    self.app.toast.error(f"Process failed at step {result.steps_completed + 1}")
            
            self.app.dispatcher.post(finish)
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _run_copy_only(self):
        """Run copy data step"""
        self.app.set_busy(True, "Copying data...")
        
        def worker():
            settings = self._get_settings()
            controller = ProcessController(settings, on_log=lambda msg, lvl: self.app.dispatcher.post(self._on_log, msg, lvl))
            result = controller.run_copy_only()
            
            def finish():
                self.app.set_busy(False)
                if result.success:
                    self.app.toast.success(f"Copied {result.rows_processed} rows")
                else:
                    self.app.toast.error(f"Copy failed: {result.message}")
            
            self.app.dispatcher.post(finish)
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _run_bjdt_only(self):
        """Run BJDT ID fetch"""
        self.app.set_busy(True, "Fetching BJDT IDs...")
        
        def worker():
            settings = self._get_settings()
            controller = ProcessController(settings, on_log=lambda msg, lvl: self.app.dispatcher.post(self._on_log, msg, lvl))
            result = controller.run_bjdt_only()
            
            def finish():
                self.app.set_busy(False)
                if result.success:
                    self.app.toast.success(f"Found {result.ids_found} IDs, {result.ids_missing} missing")
                else:
                    self.app.toast.error(f"BJDT failed: {result.message}")
            
            self.app.dispatcher.post(finish)
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _run_csv_only(self):
        """Run CSV generation"""
        self.app.set_busy(True, "Generating CSV files...")
        
        def worker():
            settings = self._get_settings()
            controller = ProcessController(settings, on_log=lambda msg, lvl: self.app.dispatcher.post(self._on_log, msg, lvl))
            result = controller.run_csv_only()
            
            def finish():
                self.app.set_busy(False)
                if result.success:
                    self.app.toast.success(f"Generated {len(result.files_created)} CSV files")
                else:
                    self.app.toast.error(f"CSV failed: {result.message}")
            
            self.app.dispatcher.post(finish)
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _run_upload_csv(self):
        """Browse and upload CSV file"""
        from tkinter import filedialog
        
        filename = filedialog.askopenfilename(
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        
        if filename:
            self._run_upload_only(filename)
    
    def _run_upload_only(self, csv_path=None):
        """Run database upload"""
        self.app.set_busy(True, "Uploading to database...")
        
        def worker():
            settings = self._get_settings()
            controller = ProcessController(settings, on_log=lambda msg, lvl: self.app.dispatcher.post(self._on_log, msg, lvl))
            result = controller.run_upload_only(csv_path)
            
            def finish():
                self.app.set_busy(False)
                if result.success:
                    self.app.toast.success(f"Uploaded {result.rows_uploaded} rows")
                elif result.duplicate:
                    self.app.toast.warning("Upload failed: Duplicate data detected")
                else:
                    self.app.toast.error(f"Upload failed: {result.message}")
            
            self.app.dispatcher.post(finish)
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _open_grid(self):
        """Open Grid Benda Uji"""
        self.app.toast.info("Grid Benda Uji not implemented yet")
        # ponytail: launch subprocess for wx grid
    
    def _open_beban_config(self):
        """Open beban config editor"""
        try:
            from modules.ui.beban_config_editor import open_config_editor
            open_config_editor(self)
        except Exception as e:
            from modules.ui.components import ErrorDialog
            ErrorDialog(self, "Error", "Failed to open config editor", detail=str(e))

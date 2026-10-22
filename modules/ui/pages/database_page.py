"""Database configuration page"""
import customtkinter as ctk
import tkinter as tk
from modules.ui.theme import SPACING, color
from modules.ui.components import (
    SectionCard, 
    LabeledEntry, 
    PasswordEntry,
    PrimaryButton,
    StatusBadge
)
import threading

class DatabasePage(ctk.CTkScrollableFrame):
    """
    Database configuration page.
    
    Replaces the old "Database configure" tab.
    """
    
    def __init__(self, parent, app, db_vars: dict):
        """
        Args:
            parent: Parent widget
            app: ModernApp instance
            db_vars: Dict of StringVars {user, password, host, port, name, table}
        """
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        self.db_vars = db_vars
        
        # Build UI
        self._build_ui()
    
    def _build_ui(self):
        """Build page content"""
        # Main card
        card = SectionCard(
            self,
            title="Database Configuration",
            subtitle="PostgreSQL connection settings"
        )
        card.pack(fill="both", expand=True)
        
        # Form container (2 columns)
        form = ctk.CTkFrame(card.body, fg_color="transparent")
        form.pack(fill="both", expand=True)
        
        # Configure grid
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        
        # Row 0: Host | Port
        LabeledEntry(
            form,
            label="Host",
            variable=self.db_vars["host"],
            width=300
        ).grid(row=0, column=0, sticky="ew", padx=(0, SPACING["md"]), pady=(0, SPACING["md"]))
        
        LabeledEntry(
            form,
            label="Port",
            variable=self.db_vars["port"],
            width=150
        ).grid(row=0, column=1, sticky="w", pady=(0, SPACING["md"]))
        
        # Row 1: User | Password
        LabeledEntry(
            form,
            label="User",
            variable=self.db_vars["user"],
            width=300
        ).grid(row=1, column=0, sticky="ew", padx=(0, SPACING["md"]), pady=(0, SPACING["md"]))
        
        PasswordEntry(
            form,
            label="Password",
            variable=self.db_vars["password"]
        ).grid(row=1, column=1, sticky="ew", pady=(0, SPACING["md"]))
        
        # Row 2: Database | Table
        LabeledEntry(
            form,
            label="Database",
            variable=self.db_vars["name"],
            width=300
        ).grid(row=2, column=0, sticky="ew", padx=(0, SPACING["md"]), pady=(0, SPACING["md"]))
        
        LabeledEntry(
            form,
            label="Table",
            variable=self.db_vars["table"],
            width=300
        ).grid(row=2, column=1, sticky="ew", pady=(0, SPACING["md"]))
        
        # Test connection section
        test_frame = ctk.CTkFrame(card.body, fg_color="transparent")
        test_frame.pack(fill="x", pady=(SPACING["lg"], 0))
        
        # Test button
        self.test_btn = PrimaryButton(
            test_frame,
            text="Test Connection",
            command=self._test_connection,
            width=150
        )
        self.test_btn.pack(side="left")
        
        # Status badge (hidden initially)
        self.status_badge = StatusBadge(test_frame, text="", state="idle")
        self.status_badge.pack(side="left", padx=(SPACING["md"], 0))
        self.status_badge.pack_forget()  # Hide initially
        
        # Result label
        self.result_label = ctk.CTkLabel(
            test_frame,
            text="",
            font=("Inter", 11),
            text_color=color("muted")
        )
        self.result_label.pack(side="left", padx=(SPACING["sm"], 0))
    
    def _test_connection(self):
        """Test database connection"""
        self.test_btn.configure(state="disabled", text="Testing...")
        self.status_badge.pack(side="left", padx=(SPACING["md"], 0))
        self.status_badge.set_state("running", "Testing")
        self.result_label.configure(text="")
        
        # Run test in background
        def test_worker():
            try:
                from sqlalchemy import create_engine, text
                import time
                
                start = time.time()
                
                # Build connection URL
                user = self.db_vars["user"].get()
                password = self.db_vars["password"].get()
                host = self.db_vars["host"].get()
                port = self.db_vars["port"].get()
                dbname = self.db_vars["name"].get()
                
                url = f"postgresql://{user}:{password}@{host}:{port}/{dbname}"
                
                # Test connection
                engine = create_engine(url, pool_pre_ping=True, connect_args={"connect_timeout": 5})
                with engine.connect() as conn:
                    conn.execute(text("SELECT 1"))
                
                elapsed_ms = int((time.time() - start) * 1000)
                
                # Success
                def update_success():
                    self.status_badge.set_state("success", "Connected")
                    self.result_label.configure(
                        text=f"✓ {elapsed_ms} ms",
                        text_color=color("success")
                    )
                    self.test_btn.configure(state="normal", text="Test Connection")
                    self.app.toast.success(f"Database connected ({elapsed_ms} ms)")
                
                self.app.dispatcher.post(update_success)
                
            except Exception as e:
                error_msg = str(e)
                
                # Shorten common errors
                if "could not connect" in error_msg or "refused" in error_msg:
                    short_msg = "Connection refused"
                elif "password authentication failed" in error_msg:
                    short_msg = "Authentication failed"
                elif "does not exist" in error_msg:
                    short_msg = "Database not found"
                elif "timeout" in error_msg.lower():
                    short_msg = "Connection timeout"
                else:
                    short_msg = error_msg[:50]
                
                # Error
                def update_error():
                    self.status_badge.set_state("error", "Failed")
                    self.result_label.configure(
                        text=f"✗ {short_msg}",
                        text_color=color("danger")
                    )
                    self.test_btn.configure(state="normal", text="Test Connection")
                
                self.app.dispatcher.post(update_error)
        
        thread = threading.Thread(target=test_worker, daemon=True)
        thread.start()

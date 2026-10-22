"""Rencana Benda Uji page"""
import customtkinter as ctk
from modules.ui.theme import SPACING, color
from modules.ui.components import SectionCard, ToggleButton, StatusBadge

class RencanaPage(ctk.CTkFrame):
    """
    Input Rencana Benda Uji page.
    
    Replaces the old "Input Rencana Benda Uji" tab.
    """
    
    def __init__(self, parent, app):
        """
        Args:
            parent: Parent widget
            app: ModernApp instance
        """
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        self.running = False
        
        self._build_ui()
    
    def _build_ui(self):
        """Build page content"""
        # Main card
        card = SectionCard(
            self,
            title="Input Rencana Benda Uji",
            subtitle="Automated web form filling with Playwright"
        )
        card.pack(fill="both", expand=True, padx=SPACING["xl"], pady=SPACING["xl"])
        
        # Center container
        center = ctk.CTkFrame(card.body, fg_color="transparent")
        center.pack(expand=True)
        
        # Status badge
        self.status_badge = StatusBadge(center, text="Ready", state="idle")
        self.status_badge.pack(pady=(0, SPACING["xl"]))
        
        # Toggle button
        self.toggle_btn = ToggleButton(
            center,
            start_text="▶ Start Process",
            stop_text="⬛ Stop Process",
            command=self._on_toggle
        )
        self.toggle_btn.pack(pady=(0, SPACING["xl"]))
        
        # Stats frame
        stats_frame = ctk.CTkFrame(center, fg_color=color("surface_alt"),
                                  corner_radius=8)
        stats_frame.pack(fill="x")
        
        # 3 stat tiles
        self.stats = {}
        for label in ["Success", "Failed", "Skipped"]:
            stat = self._create_stat_tile(stats_frame, label, "0")
            stat.pack(side="left", expand=True, padx=SPACING["md"], pady=SPACING["md"])
            self.stats[label.lower()] = stat
    
    def _create_stat_tile(self, parent, label: str, value: str):
        """Create a stat tile"""
        tile = ctk.CTkFrame(parent, fg_color="transparent")
        
        value_label = ctk.CTkLabel(
            tile,
            text=value,
            font=("Inter", 24, "bold")
        )
        value_label.pack()
        
        label_label = ctk.CTkLabel(
            tile,
            text=label,
            font=("Inter", 11),
            text_color=color("muted")
        )
        label_label.pack()
        
        # Store value label for updates
        tile.value_label = value_label
        
        return tile
    
    def _on_toggle(self, running: bool):
        """Handle start/stop toggle"""
        if running:
            self.status_badge.set_state("running", "Processing")
            self.app.toast.info("Rencana process started (not implemented)")
            # ponytail: start actual process thread
        else:
            self.status_badge.set_state("idle", "Stopped")
            self.app.toast.warning("Rencana process stopped")
            # ponytail: stop process thread
    
    def update_stats(self, success: int, failed: int, skipped: int):
        """Update stat counters (thread-safe via dispatcher)"""
        def update():
            self.stats["success"].value_label.configure(text=str(success))
            self.stats["failed"].value_label.configure(text=str(failed))
            self.stats["skipped"].value_label.configure(text=str(skipped))
        
        self.app.dispatcher.post(update)

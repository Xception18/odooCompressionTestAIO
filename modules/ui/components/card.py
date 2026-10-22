"""Reusable card component"""
import customtkinter as ctk
from modules.ui.theme import color, font, RADIUS, SPACING

class SectionCard(ctk.CTkFrame):
    """
    Card component with title, optional subtitle, and content area.
    
    Usage:
        card = SectionCard(parent, title="Database Config", subtitle="Connection settings")
        label = ctk.CTkLabel(card.body, text="Content goes here")
        label.pack()
    """
    
    def __init__(self, master, title: str, subtitle: str = None, icon=None, **kwargs):
        """
        Args:
            master: Parent widget
            title: Card title
            subtitle: Optional subtitle text
            icon: Optional CTkImage icon
            **kwargs: Additional CTkFrame arguments
        """
        # Override defaults with card styling
        kwargs.setdefault("fg_color", color("surface"))
        kwargs.setdefault("corner_radius", RADIUS["card"])
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", color("border"))
        
        super().__init__(master, **kwargs)
        
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=SPACING["lg"], pady=(SPACING["lg"], SPACING["sm"]))
        
        # Title with optional icon
        title_label = ctk.CTkLabel(
            header,
            text=title,
            font=font("section_title"),
            image=icon,
            compound="left" if icon else "none"
        )
        title_label.pack(side="left")
        
        # Subtitle
        if subtitle:
            subtitle_label = ctk.CTkLabel(
                self,
                text=subtitle,
                font=font("helper"),
                text_color=color("muted")
            )
            subtitle_label.pack(anchor="w", padx=SPACING["lg"], pady=(0, SPACING["sm"]))
        
        # Content body
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=SPACING["lg"], 
                       pady=(SPACING["sm"], SPACING["lg"]))


class StatusBadge(ctk.CTkFrame):
    """
    Small colored badge for status indicators.
    
    States: idle, running, success, warning, error
    """
    
    def __init__(self, master, text: str = "", state: str = "idle", **kwargs):
        """
        Args:
            master: Parent widget
            text: Badge text
            state: idle/running/success/warning/error
        """
        kwargs.setdefault("corner_radius", RADIUS["badge"])
        kwargs.setdefault("height", 24)
        
        super().__init__(master, **kwargs)
        
        self.label = ctk.CTkLabel(
            self,
            text=text,
            font=font("helper"),
            fg_color="transparent"
        )
        self.label.pack(padx=SPACING["sm"], pady=2)
        
        self.set_state(state, text if text else None)
    
    def set_state(self, state: str, text: str = None):
        """Update badge state and optional text"""
        colors = {
            "idle": color("muted"),
            "running": color("primary"),
            "success": color("success"),
            "warning": color("warning"),
            "error": color("danger")
        }
        
        self.configure(fg_color=colors.get(state, colors["idle"]))
        
        if text:
            self.label.configure(text=text)

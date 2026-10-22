"""Log console with color-coded levels"""
import customtkinter as ctk
from modules.ui.theme import SPACING, RADIUS

class LogConsole(ctk.CTkTextbox):
    """
    Log console with colored levels and line capping.
    
    Replaces ScrolledText with proper theming and overflow handling.
    """
    
    MAX_LINES = 5000
    
    def __init__(self, master, **kwargs):
        """
        Args:
            master: Parent widget
            **kwargs: Additional CTkTextbox arguments
        """
        # Set defaults
        kwargs.setdefault("font", ("Cascadia Mono", 12))
        kwargs.setdefault("wrap", "none")
        kwargs.setdefault("corner_radius", RADIUS["input"])
        
        super().__init__(master, **kwargs)
        
        # Configure level colors
        # Note: CTkTextbox wraps tk.Text, so tag_config works for colors only
        self.tag_config("INFO", foreground="#C9D1D9")
        self.tag_config("WARNING", foreground="#F5A524")
        self.tag_config("ERROR", foreground="#E5484D")
        self.tag_config("SUCCESS", foreground="#2EB872")
        self.tag_config("DEBUG", foreground="#8B94A1")
        
        # State
        self._autoscroll = True
        self.configure(state="disabled")
    
    def append(self, text: str, level: str = "INFO"):
        """
        Append log line with level coloring.
        
        Args:
            text: Log text (without newline)
            level: INFO/WARNING/ERROR/SUCCESS/DEBUG
        """
        self.configure(state="normal")
        
        # Add newline if text doesn't end with one
        if not text.endswith("\n"):
            text += "\n"
        
        # Insert with tag
        self.insert("end", text, level)
        
        # Cap lines
        current_line = int(self.index("end-1c").split(".")[0])
        if current_line > self.MAX_LINES:
            # Delete first 500 lines
            self.delete("1.0", "500.0")
        
        self.configure(state="disabled")
        
        # Auto-scroll to bottom
        if self._autoscroll:
            self.see("end")
    
    def set_autoscroll(self, enabled: bool):
        """Enable/disable auto-scroll"""
        self._autoscroll = enabled
    
    def clear(self):
        """Clear all logs"""
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")
    
    def get_all_text(self) -> str:
        """Get all log text"""
        return self.get("1.0", "end-1c")
    
    def filter_by_level(self, levels: list[str]):
        """
        Show only specific levels (not implemented - would need full rewrite).
        
        For now, just a placeholder. Real implementation would need to:
        1. Store all log entries with metadata
        2. Rebuild visible text on filter change
        """
        pass  # ponytail: level filtering, add when needed

"""Button components with variants"""
import customtkinter as ctk
from modules.ui.theme import color, font, RADIUS, SIZES

class PrimaryButton(ctk.CTkButton):
    """Primary action button (blue)"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault("height", SIZES["primary_button_height"])
        kwargs.setdefault("corner_radius", RADIUS["button"])
        kwargs.setdefault("font", font("body"))
        kwargs.setdefault("fg_color", color("primary"))
        kwargs.setdefault("hover_color", color("primary_hover"))
        
        super().__init__(master, **kwargs)


class SecondaryButton(ctk.CTkButton):
    """Secondary action button (outlined)"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault("height", SIZES["button_height"])
        kwargs.setdefault("corner_radius", RADIUS["button"])
        kwargs.setdefault("font", font("body"))
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", color("border"))
        kwargs.setdefault("text_color", color("text"))
        kwargs.setdefault("hover_color", color("surface_alt"))
        
        super().__init__(master, **kwargs)


class DangerButton(ctk.CTkButton):
    """Destructive action button (red)"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault("height", SIZES["button_height"])
        kwargs.setdefault("corner_radius", RADIUS["button"])
        kwargs.setdefault("font", font("body"))
        kwargs.setdefault("fg_color", color("danger"))
        kwargs.setdefault("hover_color", color("danger_hover"))
        
        super().__init__(master, **kwargs)


class SuccessButton(ctk.CTkButton):
    """Success/positive action button (green)"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault("height", SIZES["button_height"])
        kwargs.setdefault("corner_radius", RADIUS["button"])
        kwargs.setdefault("font", font("body"))
        kwargs.setdefault("fg_color", color("success"))
        kwargs.setdefault("hover_color", color("success_hover"))
        
        super().__init__(master, **kwargs)


class ToggleButton(ctk.CTkButton):
    """
    Button that toggles between two states (e.g., Start/Stop).
    
    Usage:
        btn = ToggleButton(parent, 
                          start_text="Start Process",
                          stop_text="Stop Process",
                          command=on_toggle)
        btn.set_running(True)  # Changes to stop state
    """
    
    def __init__(self, master, start_text: str = "Start", stop_text: str = "Stop",
                 command=None, **kwargs):
        """
        Args:
            master: Parent widget
            start_text: Text when idle
            stop_text: Text when running
            command: Callback(is_running: bool)
        """
        self._start_text = start_text
        self._stop_text = stop_text
        self._running = False
        self._user_command = command
        
        kwargs.setdefault("height", SIZES["primary_button_height"])
        kwargs.setdefault("corner_radius", RADIUS["button"])
        kwargs.setdefault("font", font("body"))
        kwargs["text"] = start_text
        kwargs["command"] = self._on_click
        
        # Start state colors
        kwargs.setdefault("fg_color", color("success"))
        kwargs.setdefault("hover_color", color("success_hover"))
        
        super().__init__(master, **kwargs)
        
        # Store original colors
        self._start_fg = color("success")
        self._start_hover = color("success_hover")
        self._stop_fg = color("danger")
        self._stop_hover = color("danger_hover")
    
    def _on_click(self):
        """Internal click handler"""
        self._running = not self._running
        self._update_appearance()
        
        if self._user_command:
            self._user_command(self._running)
    
    def set_running(self, running: bool):
        """Set running state programmatically"""
        self._running = running
        self._update_appearance()
    
    def is_running(self) -> bool:
        """Get current state"""
        return self._running
    
    def _update_appearance(self):
        """Update button appearance based on state"""
        if self._running:
            self.configure(
                text=self._stop_text,
                fg_color=self._stop_fg,
                hover_color=self._stop_hover
            )
        else:
            self.configure(
                text=self._start_text,
                fg_color=self._start_fg,
                hover_color=self._start_hover
            )

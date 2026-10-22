"""Design tokens and theme management for the UI"""
import customtkinter as ctk
from pathlib import Path
from typing import Literal

# Color tokens
TOKENS = {
    "dark": {
        "bg": "#121417",
        "surface": "#1B1F24",
        "surface_alt": "#232830",
        "border": "#2E343D",
        "text": "#E6E9EE",
        "muted": "#8B94A1",
        "primary": "#4A90E2",
        "primary_hover": "#3A7BC8",
        "success": "#2EB872",
        "success_hover": "#25A563",
        "danger": "#E5484D",
        "danger_hover": "#D13438",
        "warning": "#F5A524",
        "warning_hover": "#D98A00",
    },
    "light": {
        "bg": "#F4F6F9",
        "surface": "#FFFFFF",
        "surface_alt": "#EEF1F5",
        "border": "#D9DEE5",
        "text": "#1B1F24",
        "muted": "#667085",
        "primary": "#2F7BD6",
        "primary_hover": "#2467B8",
        "success": "#1E9E5C",
        "success_hover": "#1A8A51",
        "danger": "#D13438",
        "danger_hover": "#B02A2E",
        "warning": "#D98A00",
        "warning_hover": "#B87600",
    }
}

# Spacing tokens (4pt grid)
SPACING = {
    "xs": 4,
    "sm": 8,
    "md": 12,
    "lg": 16,
    "xl": 24,
    "xxl": 32
}

# Typography
FONTS = {
    "page_title": ("Inter", 24, "bold"),
    "section_title": ("Inter", 16, "bold"),
    "body": ("Inter", 13, "normal"),
    "helper": ("Inter", 11, "normal"),
    "mono": ("Cascadia Mono", 12, "normal"),
}

# Sizing
SIZES = {
    "input_height": 36,
    "button_height": 38,
    "primary_button_height": 44,
    "sidebar_width": 220,
    "sidebar_collapsed": 64,
}

# Radius
RADIUS = {
    "card": 12,
    "input": 8,
    "button": 8,
    "badge": 999,
}

def color(name: str) -> tuple[str, str]:
    """
    Return (light, dark) color tuple for CustomTkinter automatic theme switching.
    
    Args:
        name: Token name (bg, surface, text, primary, success, danger, etc.)
    
    Returns:
        Tuple of (light_color, dark_color)
    """
    return (TOKENS["light"][name], TOKENS["dark"][name])

def get_color(name: str, mode: Literal["light", "dark"] = "dark") -> str:
    """
    Get single color for a specific mode.
    
    Args:
        name: Token name
        mode: "light" or "dark"
    
    Returns:
        Hex color string
    """
    return TOKENS[mode][name]

def font(name: str) -> ctk.CTkFont:
    """
    Create a CTkFont from a font token.
    
    Args:
        name: Font token (page_title, section_title, body, helper, mono)
    
    Returns:
        CTkFont instance
    """
    family, size, weight = FONTS[name]
    return ctk.CTkFont(family=family, size=size, weight=weight)

def apply_theme(mode: Literal["dark", "light", "system"] = "dark") -> None:
    """
    Apply theme to the application.
    
    Args:
        mode: "dark", "light", or "system"
    """
    ctk.set_appearance_mode(mode)
    
    # Load custom theme JSON if it exists
    theme_path = Path(__file__).parent / "theme.json"
    if theme_path.exists():
        ctk.set_default_color_theme(str(theme_path))

def get_current_mode() -> Literal["light", "dark"]:
    """Get current appearance mode (resolves 'system' to actual mode)"""
    mode = ctk.get_appearance_mode()
    return "dark" if mode == "Dark" else "light"

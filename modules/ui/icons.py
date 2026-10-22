"""Icon loading and caching for the UI"""
import customtkinter as ctk
from pathlib import Path
from typing import Optional
from PIL import Image

# Icon cache
_icon_cache = {}

def load_icon(name: str, size: tuple[int, int] = (20, 20)) -> Optional[ctk.CTkImage]:
    """
    Load an icon with light/dark variants.
    
    Args:
        name: Icon name (without extension or _light/_dark suffix)
        size: (width, height) tuple
    
    Returns:
        CTkImage with light/dark variants, or None if not found
    
    Icon file structure:
        assets/icons/{name}_light.png
        assets/icons/{name}_dark.png
    """
    cache_key = (name, size)
    
    if cache_key in _icon_cache:
        return _icon_cache[cache_key]
    
    icons_dir = Path(__file__).parent.parent.parent / "assets" / "icons"
    light_path = icons_dir / f"{name}_light.png"
    dark_path = icons_dir / f"{name}_dark.png"
    
    # If neither exists, return None
    if not light_path.exists() and not dark_path.exists():
        return None
    
    # Use same image for both if only one exists
    if not light_path.exists():
        light_path = dark_path
    if not dark_path.exists():
        dark_path = light_path
    
    try:
        light_img = Image.open(light_path)
        dark_img = Image.open(dark_path)
        
        icon = ctk.CTkImage(
            light_image=light_img,
            dark_image=dark_img,
            size=size
        )
        
        _icon_cache[cache_key] = icon
        return icon
    
    except Exception as e:
        print(f"[Icons] Failed to load {name}: {e}")
        return None

def clear_icon_cache():
    """Clear the icon cache (useful for theme changes)"""
    global _icon_cache
    _icon_cache = {}

# Placeholder icon generator (for development before real icons exist)
def placeholder_icon(size: tuple[int, int] = (20, 20)) -> ctk.CTkImage:
    """Create a placeholder icon (gray square)"""
    cache_key = ("_placeholder", size)
    
    if cache_key in _icon_cache:
        return _icon_cache[cache_key]
    
    # Create simple colored square
    from PIL import Image, ImageDraw
    
    img_light = Image.new('RGBA', size, (200, 200, 200, 255))
    img_dark = Image.new('RGBA', size, (100, 100, 100, 255))
    
    # Draw simple shape
    draw_light = ImageDraw.Draw(img_light)
    draw_dark = ImageDraw.Draw(img_dark)
    
    margin = 4
    draw_light.ellipse([margin, margin, size[0]-margin, size[1]-margin], 
                       fill=(150, 150, 150, 255))
    draw_dark.ellipse([margin, margin, size[0]-margin, size[1]-margin], 
                      fill=(150, 150, 150, 255))
    
    icon = ctk.CTkImage(light_image=img_light, dark_image=img_dark, size=size)
    _icon_cache[cache_key] = icon
    return icon

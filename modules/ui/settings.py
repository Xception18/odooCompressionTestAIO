"""Settings persistence"""
import json
from pathlib import Path
from typing import Optional
from dataclasses import dataclass, asdict

@dataclass
class AppSettings:
    """Application settings"""
    # Files
    rencana_path: str = "Rencana/RENCANA.xlsx"
    pengujian_path: str = "Pengujian/PENGUJIAN.xlsx"
    output_dir: str = "Output"
    
    # Sheets
    sheet_source: str = "Report Excel"
    sheet_odoo: str = "ODOO"
    sheet_pengujian: str = "pengujian"
    sheet_bjdt: str = "bjdt_id"
    
    # Database
    db_host: str = "localhost"
    db_port: str = "5432"
    db_user: str = "postgres"
    db_name: str = "alatujidb"
    db_table: str = "pengujian"
    
    # UI
    theme: str = "dark"  # dark/light
    window_width: int = 1200
    window_height: int = 800


def get_settings_path() -> Path:
    """Get settings file path"""
    return Path.home() / ".hermes" / "excel_processor_settings.json"


def load_settings() -> AppSettings:
    """Load settings from file"""
    settings_path = get_settings_path()
    
    if settings_path.exists():
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return AppSettings(**data)
        except Exception as e:
            print(f"[Settings] Failed to load: {e}, using defaults")
    
    return AppSettings()


def save_settings(settings: AppSettings) -> bool:
    """Save settings to file"""
    settings_path = get_settings_path()
    
    try:
        # Create directory
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save JSON
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(asdict(settings), f, indent=2)
        
        return True
    except Exception as e:
        print(f"[Settings] Failed to save: {e}")
        return False

"""Configuration manager for beban calculation logic"""
import json
import logging
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger(__name__)

class BebanConfig:
    """Manages beban calculation configuration"""
    
    DEFAULT_CONFIG_PATH = Path(__file__).parent / "beban_config.json"
    
    def __init__(self, config_path: Optional[str] = None):
        self.config_path = config_path or self.DEFAULT_CONFIG_PATH
        self.config: Dict = {}
        self.load()
    
    def load(self) -> bool:
        """Load config from JSON file"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            logger.info(f"Loaded config from {self.config_path}")
            return True
        except FileNotFoundError:
            logger.warning(f"Config file not found: {self.config_path}")
            self.config = {}
            return False
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in config: {e}")
            self.config = {}
            return False
    
    def save(self, path: Optional[str] = None) -> bool:
        """Save config to JSON file"""
        target_path = path or self.config_path
        try:
            with open(target_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logger.info(f"Saved config to {target_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
            return False
    
    def get_mutu_config(self, mutu: str) -> Optional[Dict]:
        """Get config for specific mutu"""
        # Try exact match first
        if mutu in self.config:
            return self.config[mutu]
        
        # Try partial match (e.g., "Class A-2" in "Class A-2 NFA")
        for key in self.config.keys():
            if key in mutu:
                return self.config[key]
        
        return None
    
    def get_all_mutu_names(self) -> list:
        """Get list of all configured mutu names"""
        return list(self.config.keys())
    
    def add_mutu(self, mutu_name: str, config: Dict) -> bool:
        """Add or update mutu configuration"""
        self.config[mutu_name] = config
        return self.save()
    
    def delete_mutu(self, mutu_name: str) -> bool:
        """Delete mutu configuration"""
        if mutu_name in self.config:
            del self.config[mutu_name]
            return self.save()
        return False

# Global instance
_config_instance: Optional[BebanConfig] = None

def get_beban_config() -> BebanConfig:
    """Get global BebanConfig instance"""
    global _config_instance
    if _config_instance is None:
        _config_instance = BebanConfig()
    return _config_instance

"""Config-driven beban calculation logic (extracted from excel_handler.py)"""
import random
from typing import Optional, Dict

def calculate_beban_from_config(
    mutu: str,
    umur: int,
    docket: str,
    kontraktor: str,
    config: Dict,
    previous_docket: Optional[str],
    previous_beban: Optional[float],
    previous_mutu: Optional[str]
) -> tuple[Optional[str], Optional[str], Optional[float]]:
    """
    Calculate beban using config dict
    
    Returns:
        (beban_value, updated_docket, updated_beban)
    """
    
    if umur not in [7, 28]:
        return "0.00", previous_docket, previous_beban
    
    # Umur 7 - simple range
    if umur == 7:
        u7_config = config.get("umur_7", {})
        min_val = u7_config.get("min", 0)
        max_val = u7_config.get("max", 0)
        
        if min_val == 0 and max_val == 0:
            return "0.00", previous_docket, previous_beban
        
        random_beban = random.uniform(min_val, max_val)
        return f"{random_beban:.2f}", previous_docket, previous_beban
    
    # Umur 28 - complex logic
    u28_config = config.get("umur_28", {})
    if not u28_config:
        return "0.00", previous_docket, previous_beban
    
    initial = u28_config.get("initial", {})
    mean = u28_config.get("mean", 0)
    same_up = u28_config.get("same_docket_adjust_up", {})
    same_down = u28_config.get("same_docket_adjust_down", {})
    diff_up = u28_config.get("diff_docket_adjust_up", {})
    diff_down = u28_config.get("diff_docket_adjust_down", {})
    range_min = u28_config.get("range_min", 0)
    range_max = u28_config.get("range_max", 0)
    
    new_beban = 0.0
    
    # Scenario 1: Initial condition (no previous data)
    if (previous_docket is None or previous_docket == 0.0 or
        previous_beban is None or previous_beban == 0.0):
        new_beban = random.uniform(initial.get("min", 0), initial.get("max", 0))
    
    # Scenario 2: Same docket, beban exists
    elif previous_docket == docket and previous_beban is not None:
        if previous_beban >= mean:
            adjustment = random.uniform(same_down.get("min", 0), same_down.get("max", 0))
        else:
            adjustment = random.uniform(same_up.get("min", 0), same_up.get("max", 0))
        new_beban = max(range_min, min(previous_beban + adjustment, range_max))
    
    # Scenario 3: Different docket, same mutu
    elif previous_docket != docket and previous_mutu == mutu:
        if previous_beban >= mean:
            adjustment = random.uniform(diff_down.get("min", 0), diff_down.get("max", 0))
        else:
            adjustment = random.uniform(diff_up.get("min", 0), diff_up.get("max", 0))
        new_beban = previous_beban + adjustment
        new_beban = max(range_min, min(new_beban, range_max))
    
    # Scenario 4: Different docket, different mutu
    elif previous_docket != docket and previous_mutu != mutu:
        new_beban = random.uniform(initial.get("min", 0), initial.get("max", 0))
    
    else:
        # Fallback
        new_beban = random.uniform(initial.get("min", 0), initial.get("max", 0))
    
    # Update tracking
    updated_docket = str(docket) if docket is not None else None
    updated_beban = new_beban
    
    return f"{new_beban:.2f}", updated_docket, updated_beban

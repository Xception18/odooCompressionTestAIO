import random
from datetime import datetime, timedelta

def generate_random_slump_test(slump_rencana):
    """Generate random slump test value (±2 from slump_rencana)"""
    try:
        base_value = float(slump_rencana)
        
        # Special case for slump 55
        if base_value == 55:
            result = 60 + random.uniform(-5, 5)
        else:
            result = base_value + random.uniform(-1, 2)
        
        return str(int(round(result)))  # Convert to integer to remove decimal
    except:
        result = random.uniform(11, 14)
        return str(int(round(result)))

def generate_random_yield():
    """Generate random yield value between 0.97-0.99"""
    return str(round(random.uniform(0.97, 0.99), 2))

def calculate_jam_sample(base_time):
    """Calculate jam sample by adding 1:10 to 1:50 hours to base time"""
    try:
        additional_minutes = random.randint(65, 90)
        str_val = str(base_time).strip()
        
        # Check if base_time contains date (format: 'DD/MM/YYYY HH:MM:SS' or 'YYYY-MM-DD HH:MM:SS')
        if ' ' in str_val:
            dt = None
            # Try various formats
            formats = [
                '%d/%m/%Y %H:%M:%S',
                '%d/%m/%Y %H:%M',
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d %H:%M'
            ]
            
            for fmt in formats:
                try:
                    dt = datetime.strptime(str_val, fmt)
                    break
                except ValueError:
                    continue
            
            if dt:
                new_dt = dt + timedelta(minutes=additional_minutes)
                return new_dt.strftime('%d/%m/%Y %H:%M:%S')
            else:
                # If parsing fails despite containing space, return random fallback or handle error
                # For now falling back to existing logic (which might error out or hit the except block below)
                raise ValueError("Unknown date format")
        else:
            # Original format HH:MM
            # Handle seconds if present but we only return HH:MM as per original logic for time-only
            time_part = str_val
            if time_part.count(':') == 2:
                time_part = ':'.join(time_part.split(':')[:2])
            
            base_hour, base_minute = map(int, time_part.split(':'))
            
            total_minutes = base_hour * 60 + base_minute + additional_minutes
            final_hour = (total_minutes // 60) % 24
            final_minute = total_minutes % 60
            
            return f"{final_hour:02d}:{final_minute:02d}"
    except:
        # Fallback to random time
        hour = random.randint(10, 15)
        minute = random.randint(0, 59)
        return f"{hour:02d}:{minute:02d}"

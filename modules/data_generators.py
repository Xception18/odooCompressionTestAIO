import random

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
        # Check if base_time contains date (format: 'YYYY-MM-DD HH:MM:SS' or 'YYYY-MM-DD HH:MM')
        if ' ' in str(base_time):
            # Extract time part from datetime string
            time_part = str(base_time).split(' ')[1]
            # Handle seconds if present
            if time_part.count(':') == 2:
                time_part = ':'.join(time_part.split(':')[:2])  # Keep only HH:MM
            base_hour, base_minute = map(int, time_part.split(':'))
        else:
            # Original format HH:MM
            base_hour, base_minute = map(int, str(base_time).split(':'))
        
        # Add random time between 1:5 to 1:30 hours
        additional_minutes = random.randint(65, 90)
        
        total_minutes = base_hour * 60 + base_minute + additional_minutes
        final_hour = (total_minutes // 60) % 24
        final_minute = total_minutes % 60
        
        return f"{final_hour:02d}:{final_minute:02d}"
    except:
        # Fallback to random time
        hour = random.randint(10, 15)
        minute = random.randint(0, 59)
        return f"{hour:02d}:{minute:02d}"

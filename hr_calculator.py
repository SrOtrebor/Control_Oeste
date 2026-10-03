from datetime import datetime, timedelta

def parse_time(time_str):
    try:
        return datetime.strptime(time_str, '%H:%M:%S').time()
    except:
        return None

def time_to_datetime(t, base_date):
    """Convierte un objeto time a datetime usando un base_date"""
    return datetime.combine(base_date, t)

def calcular_horas_turno(entrada_str, salida_str, aplica_nocturnidad=True):
    """
    Calcula horas diurnas y nocturnas dado una hora de entrada y salida.
    El horario nocturno es fijo: de 21:00 a 06:00.
    
    Retorna: (horas_diurnas_float, horas_nocturnas_float)
    """
    t_in = parse_time(entrada_str)
    t_out = parse_time(salida_str)
    
    if not t_in or not t_out:
        return 0.0, 0.0
        
    # Asumimos que la entrada es hoy
    base_date = datetime.today().date()
    dt_in = time_to_datetime(t_in, base_date)
    dt_out = time_to_datetime(t_out, base_date)
    
    # Si la salida es menor a la entrada, significa que cruzó la medianoche
    if dt_out < dt_in:
        dt_out += timedelta(days=1)
        
    total_seconds = (dt_out - dt_in).total_seconds()
    
    if not aplica_nocturnidad:
        # Horas planas
        return round(total_seconds / 3600.0, 2), 0.0
        
    # Definir franjas nocturnas relevantes
    # Noche 1: de ayer a 06:00 de hoy
    n1_start = time_to_datetime(datetime.strptime('21:00:00', '%H:%M:%S').time(), base_date - timedelta(days=1))
    n1_end = time_to_datetime(datetime.strptime('06:00:00', '%H:%M:%S').time(), base_date)
    
    # Noche 2: de 21:00 de hoy a 06:00 de mañana
    n2_start = time_to_datetime(datetime.strptime('21:00:00', '%H:%M:%S').time(), base_date)
    n2_end = time_to_datetime(datetime.strptime('06:00:00', '%H:%M:%S').time(), base_date + timedelta(days=1))
    
    nocturnal_seconds = 0.0
    
    # Función para calcular solapamiento entre dos rangos de datetime
    def get_overlap(start1, end1, start2, end2):
        latest_start = max(start1, start2)
        earliest_end = min(end1, end2)
        delta = (earliest_end - latest_start).total_seconds()
        return max(0.0, delta)
        
    # Sumar solapamientos
    nocturnal_seconds += get_overlap(dt_in, dt_out, n1_start, n1_end)
    nocturnal_seconds += get_overlap(dt_in, dt_out, n2_start, n2_end)
    
    diurnal_seconds = total_seconds - nocturnal_seconds
    
    return round(diurnal_seconds / 3600.0, 2), round(nocturnal_seconds / 3600.0, 2)

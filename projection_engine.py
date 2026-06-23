import sqlite3
import datetime
import urllib.request
import re
import asyncio
from calculator import calculate_liquidation, calculate_inclusive_work_days

def parse_target_date(target_date_str):
    if not target_date_str:
        return None
    import re
    target_date_str = target_date_str.strip()
    m1 = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$', target_date_str)
    if m1:
        return int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
    m2 = re.match(r'^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$', target_date_str)
    if m2:
        return int(m2.group(3)), int(m2.group(2)), int(m2.group(1))
    return None


# Default holidays for 2026 fallback
DEFAULT_HOLIDAYS_2026 = [
    "2026-01-01", # Año Nuevo
    "2026-04-03", # Viernes Santo
    "2026-04-04", # Sábado Santo
    "2026-05-01", # Día del Trabajo
    "2026-05-21", # Glorias Navales
    "2026-06-07", # Batalla de Arica
    "2026-06-21", # Pueblos Indígenas
    "2026-06-29", # San Pedro y San Pablo
    "2026-07-16", # Virgen del Carmen
    "2026-08-15", # Asunción de la Virgen
    "2026-08-20", # Nacimiento O'Higgins
    "2026-09-18", # Fiestas Patrias
    "2026-09-19", # Glorias del Ejército
    "2026-10-12", # Encuentro de Dos Mundos
    "2026-10-31", # Iglesias Evangélicas
    "2026-11-01", # Todos los Santos
    "2026-12-08", # Inmaculada Concepción
    "2026-12-25", # Navidad
]

_FERIADOS_CACHE = {}

async def get_feriados(year: int = 2026) -> list[str]:
    """
    Asynchronously fetches holidays from www.feriados.cl and normalizes them.
    Falls back to a hardcoded 2026 list on failure. Caches results in memory.
    """
    global _FERIADOS_CACHE
    try:
        year = int(year)
    except Exception:
        pass
    print(f"DEBUG get_feriados: year={year}, type={type(year)}, cache={list(_FERIADOS_CACHE.keys())}")
    if year in _FERIADOS_CACHE:
        return _FERIADOS_CACHE[year]
    if year == 2026:
        _FERIADOS_CACHE[year] = DEFAULT_HOLIDAYS_2026
        return DEFAULT_HOLIDAYS_2026
    def fetch():
        try:
            req = urllib.request.Request(
                "https://www.feriados.cl",
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                return response.read().decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"Error fetching from feriados.cl: {e}")
            return None

    html = await asyncio.to_thread(fetch)
    if not html:
        return DEFAULT_HOLIDAYS_2026

    # Match tables/cells in feriados.cl format
    matches = re.findall(r'<td>([^<]+,\s+\d+\s+de\s+[^<]+)</td>', html)
    if not matches:
        return DEFAULT_HOLIDAYS_2026

    MONTHS_MAP = {
        "enero": "01",
        "febrero": "02",
        "marzo": "03",
        "abril": "04",
        "mayo": "05",
        "junio": "06",
        "julio": "07",
        "agosto": "08",
        "septiembre": "09",
        "octubre": "10",
        "noviembre": "11",
        "diciembre": "12"
    }

    holidays = []
    for m in matches:
        parts = m.split(",")
        if len(parts) < 2:
            continue
        day_month = parts[1].strip()
        dm_parts = day_month.split(" de ")
        if len(dm_parts) < 2:
            continue
        day_str = dm_parts[0].strip()
        month_name = dm_parts[1].strip().lower()
        month_str = MONTHS_MAP.get(month_name)
        if not month_str:
            continue
        try:
            day_int = int(day_str)
            date_iso = f"{year}-{month_str}-{day_int:02d}"
            if date_iso not in holidays:
                holidays.append(date_iso)
        except ValueError:
            continue

    res = holidays if holidays else DEFAULT_HOLIDAYS_2026
    _FERIADOS_CACHE[year] = res
    return res

def calculate_weekday_holidays(year: int, month: int, holidays: list[str]) -> int:
    """
    Returns the number of holidays that fall on a weekday (Monday to Saturday, i.e., index < 6).
    """
    count = 0
    for h in holidays:
        try:
            dt = datetime.datetime.strptime(h, "%Y-%m-%d").date()
            if dt.year == year and dt.month == month:
                if dt.weekday() < 6: # Monday through Saturday
                    count += 1
        except Exception:
            continue
    return count

def project_obra_payroll(conn, id_obra, period_origin: str, year: int, month: int, snapshot_tag: str = None, overrides: dict = None, target_date: str = None) -> dict:
    """
    Projects the next month's payroll for active (non-finiquitado) employees of a given obra (or list of obras).
    - Resolves active employees (excluding those finiquitados).
    - Group by RUT strictly to average the last 12 months of variable income/deductions.
    - Adjust base days worked by subtracting weekday holidays.
    - Supports in-memory overrides to simulate finiquitos.
    """
    cursor = conn.cursor()
    
    if isinstance(id_obra, list):
        id_obras = id_obra
    elif id_obra:
        id_obras = [id_obra]
    else:
        id_obras = []
        
    if not id_obras:
        # If no obras are specified, return empty structure early
        proj_period = f"{year:04d}-{month:02d}"
        return {
            "id_obra": id_obra,
            "periodo_proyectado": proj_period,
            "base_dias_trabajados": 30,
            "feriados_habiles": 0,
            "count_empleados_proyectados": 0,
            "total_costo_empresa": 0,
            "total_imponible": 0,
            "total_haberes": 0,
            "total_liquido": 0,
            "empleados": [],
            "simulated_finiquitos": [],
            "total_finiquitos_monto": 0
        }
        
    placeholders = ",".join("?" for _ in id_obras)
    
    # 1. Resolve active employees list
    if snapshot_tag:
        query_snapshot = f"""
            SELECT * FROM empleados_snapshots 
            WHERE snapshot_tag = ? AND id_obra IN ({placeholders})
        """
        cursor.execute(query_snapshot, [snapshot_tag] + id_obras)
        employees_source = [dict(row) for row in cursor.fetchall()]
        
        # Keep only employees who were active at snapshot time
        active_snapshot_employees = [
            e for e in employees_source 
            if e["fecha_finiquito"] is None or e["fecha_finiquito"] == ""
        ]
        
        # Get currently active RUTs in main table
        query_active = f"""
            SELECT DISTINCT rut FROM empleados 
            WHERE id_obra IN ({placeholders}) AND (fecha_finiquito IS NULL OR fecha_finiquito = '')
        """
        cursor.execute(query_active, id_obras)
        currently_active_ruts = {r[0] for r in cursor.fetchall()}
        
        # Keep snapshot employees whose RUT is still active in the main table
        active_employees = [e for e in active_snapshot_employees if e["rut"] in currently_active_ruts]
    else:
        query_employees = f"""
            SELECT * FROM empleados 
            WHERE id_obra IN ({placeholders}) AND (fecha_finiquito IS NULL OR fecha_finiquito = '')
        """
        cursor.execute(query_employees, id_obras)
        active_employees = [dict(row) for row in cursor.fetchall()]
        
    # Calculate target range of periods
    parsed_target = parse_target_date(target_date) if target_date else None
    if parsed_target:
        t_yr, t_mo, t_dy = parsed_target
    else:
        t_yr = int(year)
        t_mo = int(month)
        import calendar
        t_dy = calendar.monthrange(t_yr, t_mo)[1]
        target_date = f"{t_yr:04d}-{t_mo:02d}-{t_dy:02d}"

    try:
        orig_yr = int(period_origin[:4])
        orig_mo = int(period_origin[5:7])
    except Exception:
        orig_yr, orig_mo = 2026, 5
        
    periods_in_range = []
    curr_yr = orig_yr
    curr_mo = orig_mo
    while True:
        curr_mo += 1
        if curr_mo > 12:
            curr_mo = 1
            curr_yr += 1
        if curr_yr > t_yr or (curr_yr == t_yr and curr_mo > t_mo):
            break
        periods_in_range.append((curr_yr, curr_mo))
        
    if not periods_in_range:
        periods_in_range.append((t_yr, t_mo))
        
    # 2. Adjust base days based on holidays
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            holidays = DEFAULT_HOLIDAYS_2026
        else:
            holidays = asyncio.run(get_feriados(t_yr))
    except Exception:
        holidays = DEFAULT_HOLIDAYS_2026

    # Calculate days in each month of the range using real business days (Mon-Sat, excluding holidays)
    total_days = 0
    num_months = 0.0
    import calendar
    for y, m in periods_in_range:
        first_day_of_month = datetime.date(y, m, 1)
        last_day_of_month = datetime.date(y, m, calendar.monthrange(y, m)[1])
        busdays_in_month = calculate_inclusive_work_days(first_day_of_month, last_day_of_month, holidays)
        
        if y == t_yr and m == t_mo:
            if t_dy >= last_day_of_month.day:
                days_worked = busdays_in_month
                scale_factor = 1.0
            else:
                target_dt_obj = datetime.date(t_yr, t_mo, t_dy)
                days_worked = calculate_inclusive_work_days(first_day_of_month, target_dt_obj, holidays)
                scale_factor = float(days_worked) / float(busdays_in_month) if busdays_in_month > 0 else 0.0
        else:
            days_worked = busdays_in_month
            scale_factor = 1.0
            
        total_days += days_worked
        num_months += scale_factor
        
    base_days_worked = total_days

    # Sum weekday holidays strictly from the day after the origin period to target_date
    total_weekday_holidays_count = 0
    target_dt_obj = datetime.date(t_yr, t_mo, t_dy)
    proj_start_date = datetime.date(orig_yr, orig_mo, calendar.monthrange(orig_yr, orig_mo)[1])
    for h in holidays:
        try:
            dt = datetime.datetime.strptime(h, "%Y-%m-%d").date()
            if proj_start_date < dt <= target_dt_obj:
                if dt.weekday() < 6: # Monday through Saturday
                    total_weekday_holidays_count += 1
        except Exception:
            continue
    
    # 3. Calculate start period for 12 months history
    try:
        yr = int(period_origin[:4])
        mo = int(period_origin[5:7])
    except Exception:
        yr, mo = 2026, 5
        
    start_mo = mo - 11
    start_yr = yr
    while start_mo <= 0:
        start_mo += 12
        start_yr -= 1
    start_period = f"{start_yr:04d}-{start_mo:02d}"
    
    # Parameters for projected month
    proj_period = f"{year:04d}-{month:02d}"
    params = {
        "uf": 40610.69,
        "utm": 70588.00,
        "imm": 539000.00,
        "sis_tasa": 1.62,
        "mutual_tasa": 0.93,
        "tope_imponible_afp_uf": 90.0,
        "tope_imponible_afc_uf": 135.2,
        "periodo": proj_period
    }
    
    if overrides is None:
        overrides = {}
        
    projected_records = []
    simulated_finiquitos = []
    total_costo_empresa = 0
    total_imponible = 0
    total_haberes = 0
    total_liquido = 0
    total_finiquitos_monto = 0
    
    import calculator
    
    for e in active_employees:
        rut = e["rut"]
        contrato = e["contrato"]
        
        # Check if this employee is finiquitado in overrides
        if rut in overrides and overrides[rut].get("finiquitar"):
            ov = overrides[rut]
            causal = ov.get("causal", "161")
            fecha_termino_str = ov.get("fecha_termino", f"{year}-{month:02d}-30")
            
            # Load last liq colacion/movilizacion strictly by RUT
            cursor.execute("""
                SELECT SUM(colacion) as colacion, SUM(movilizacion) as movilizacion, SUM(dias_trabajados) as dias_trabajados 
                FROM liquidaciones 
                WHERE rut = ? AND periodo = (
                    SELECT MAX(periodo) FROM liquidaciones WHERE rut = ?
                )
            """, (rut, rut))
            liq = cursor.fetchone()
            last_liq = dict(liq) if liq else None
            
            # Build inputs dictionary for the pure finiquito calculation
            inputs = {
                "aviso_previo": int(ov.get("aviso_previo", 0) or 0),
                "dias_vacaciones_override": ov.get("dias_vacaciones_override"),
                "afc_override": ov.get("afc_override"),
                "vac_progresivo": float(ov.get("vac_progresivo", 0.0) or 0.0),
                "vac_inhabiles": float(ov.get("vac_inhabiles", 0.0) or 0.0),
                "vac_tomadas": float(ov.get("vac_tomadas", 0.0) or 0.0),
                "ts_yesno": ov.get("ts_yesno", "NO"),
                "compensatoria_monto": int(ov.get("compensatoria_monto", 0) or 0),
                "prestamo_monto": int(ov.get("prestamo_monto", 0) or 0),
                "bono_1": int(ov.get("bono_1", 0) or 0),
                "bono_2": int(ov.get("bono_2", 0) or 0),
                "sueldo_base_override": ov.get("sueldo_base_override"),
                "gratificacion_override": ov.get("gratificacion_override"),
                "movilizacion_override": ov.get("movilizacion_override")
            }
            
            # Query historical AFC sum strictly by RUT
            cursor.execute("SELECT SUM(aporte_afc) FROM liquidaciones WHERE rut = ?", (rut,))
            historical_afc = cursor.fetchone()[0] or 0
            inputs["historical_afc_sum"] = historical_afc
            
            # Run calculations
            res_fin = calculator.calculate_finiquito_pure(e, params, last_liq, inputs, causal, fecha_termino_str, conn=conn)
            
            # Apply manual override of total_finiquito if present
            legal_minimum = res_fin.get("total_finiquito", 0)
            res_fin["legal_minimum"] = legal_minimum
            res_fin["warning"] = None
            
            if "override_monto" in ov and ov["override_monto"] is not None and str(ov["override_monto"]).strip() != "":
                override_val = int(ov["override_monto"])
                if override_val < legal_minimum:
                    res_fin["warning"] = f"El monto ingresado (${override_val:,} CLP) es menor al mínimo legal calculado (${legal_minimum:,} CLP)."
                res_fin["total_finiquito"] = override_val
                
            simulated_finiquitos.append(res_fin)
            total_finiquitos_monto += res_fin.get("total_finiquito", 0)
            continue
            
        # 4. Query 12 months average variables grouped strictly by RUT
        cursor.execute("""
            SELECT 
                AVG(sum_dias) as avg_dias,
                AVG(sum_horas_extras) as avg_horas_extras,
                AVG(sum_bono_descanso) as avg_bono_descanso,
                AVG(sum_bono_feriado) as avg_bono_feriado,
                AVG(sum_bono_incentivo) as avg_bono_incentivo,
                AVG(sum_bono_responsabilidad) as avg_bono_responsabilidad,
                AVG(sum_bono_gestion) as avg_bono_gestion,
                AVG(sum_bono_permanencia) as avg_bono_permanencia,
                AVG(sum_colacion) as avg_colacion,
                AVG(sum_movilizacion) as avg_movilizacion,
                AVG(sum_pasajes) as avg_pasajes,
                AVG(sum_traslados) as avg_traslados,
                AVG(sum_bono_estudios) as avg_bono_estudios,
                AVG(sum_bono_fallecimiento) as avg_bono_fallecimiento,
                AVG(sum_apvi) as avg_apvi,
                AVG(sum_anticipo) as avg_anticipo,
                AVG(sum_ccaf_credito) as avg_ccaf_credito,
                AVG(sum_ccaf_prestamo) as avg_ccaf_prestamo,
                AVG(sum_retencion_judicial) as avg_retencion_judicial,
                AVG(sum_prestamos_empresa) as avg_prestamos_empresa,
                AVG(sum_seguro_complementario) as avg_seguro_complementario,
                AVG(sum_falp) as avg_falp,
                AVG(sum_licencia_dias) as avg_licencia_dias
            FROM (
                SELECT 
                    periodo,
                    SUM(dias_trabajados) as sum_dias,
                    SUM(horas_extras) as sum_horas_extras,
                    SUM(bono_descanso) as sum_bono_descanso,
                    SUM(bono_feriado) as sum_bono_feriado,
                    SUM(bono_incentivo) as sum_bono_incentivo,
                    SUM(bono_responsabilidad) as sum_bono_responsabilidad,
                    SUM(bono_gestion) as sum_bono_gestion,
                    SUM(bono_permanencia) as sum_bono_permanencia,
                    SUM(colacion) as sum_colacion,
                    SUM(movilizacion) as sum_movilizacion,
                    SUM(pasajes) as sum_pasajes,
                    SUM(traslados) as sum_traslados,
                    SUM(bono_estudios) as sum_bono_estudios,
                    SUM(bono_fallecimiento) as sum_bono_fallecimiento,
                    SUM(descuento_apvi) as sum_apvi,
                    SUM(descuento_anticipo) as sum_anticipo,
                    SUM(descuento_ccaf_credito) as sum_ccaf_credito,
                    SUM(descuento_ccaf_prestamo) as sum_ccaf_prestamo,
                    SUM(descuento_retencion_judicial) as sum_retencion_judicial,
                    SUM(descuento_prestamos_empresa) as sum_prestamos_empresa,
                    SUM(descuento_seguro_complementario) as sum_seguro_complementario,
                    SUM(descuento_falp) as sum_falp,
                    SUM(licencia_dias) as sum_licencia_dias
                FROM liquidaciones
                WHERE rut = ? AND periodo <= ? AND periodo >= ?
                GROUP BY periodo
            )
        """, (rut, period_origin, start_period))
        avg_row = cursor.fetchone()
        
        # Determine days worked for a single month and cap by 30
        lic_days = round(avg_row["avg_licencia_dias"] or 0) if (avg_row and avg_row["avg_licencia_dias"] is not None) else 0
        dias_trabajados_single = max(0, 30 - lic_days)
        
        inputs = {
            "dias_trabajados": dias_trabajados_single,
            "licencia_dias": lic_days,
            "horas_extras_qty": round(avg_row["avg_horas_extras"] or 0) if (avg_row and avg_row["avg_horas_extras"]) else 0,
            "bono_descanso": round(avg_row["avg_bono_descanso"] or 0) if (avg_row and avg_row["avg_bono_descanso"]) else 0,
            "bono_feriado": round(avg_row["avg_bono_feriado"] or 0) if (avg_row and avg_row["avg_bono_feriado"]) else 0,
            "bono_incentivo": round(avg_row["avg_bono_incentivo"] or 0) if (avg_row and avg_row["avg_bono_incentivo"]) else 0,
            "bono_responsabilidad": round(avg_row["avg_bono_responsabilidad"] or 0) if (avg_row and avg_row["avg_bono_responsabilidad"]) else 0,
            "bono_gestion": round(avg_row["avg_bono_gestion"] or 0) if (avg_row and avg_row["avg_bono_gestion"]) else 0,
            "bono_permanencia": round(avg_row["avg_bono_permanencia"] or 0) if (avg_row and avg_row["avg_bono_permanencia"]) else 0,
            "colacion": round(avg_row["avg_colacion"] or 0) if (avg_row and avg_row["avg_colacion"]) else 0,
            "movilizacion": round(avg_row["avg_movilizacion"] or 0) if (avg_row and avg_row["avg_movilizacion"]) else 0,
            "pasajes": round(avg_row["avg_pasajes"] or 0) if (avg_row and avg_row["avg_pasajes"]) else 0,
            "traslados": round(avg_row["avg_traslados"] or 0) if (avg_row and avg_row["avg_traslados"]) else 0,
            "bono_estudios": round(avg_row["avg_bono_estudios"] or 0) if (avg_row and avg_row["avg_bono_estudios"]) else 0,
            "bono_fallecimiento": round(avg_row["avg_bono_fallecimiento"] or 0) if (avg_row and avg_row["avg_bono_fallecimiento"]) else 0,
            "apvi": round(avg_row["avg_apvi"] or 0) if (avg_row and avg_row["avg_apvi"]) else 0,
            "anticipo": round(avg_row["avg_anticipo"] or 0) if (avg_row and avg_row["avg_anticipo"]) else 0,
            "ccaf_credito": round(avg_row["avg_ccaf_credito"] or 0) if (avg_row and avg_row["avg_ccaf_credito"]) else 0,
            "ccaf_prestamo": round(avg_row["avg_ccaf_prestamo"] or 0) if (avg_row and avg_row["avg_ccaf_prestamo"]) else 0,
            "retencion_judicial": round(avg_row["avg_retencion_judicial"] or 0) if (avg_row and avg_row["avg_retencion_judicial"]) else 0,
            "prestamos_empresa": round(avg_row["avg_prestamos_empresa"] or 0) if (avg_row and avg_row["avg_prestamos_empresa"]) else 0,
            "seguro_complementario": round(avg_row["avg_seguro_complementario"] or 0) if (avg_row and avg_row["avg_seguro_complementario"]) else 0,
            "falp": round(avg_row["avg_falp"] or 0) if (avg_row and avg_row["avg_falp"]) else 0,
        }
        
        # Run Chilean payroll motor for a single month
        res_single = calculate_liquidation(e, inputs, params)
        
        # Scale the single month results by num_months
        res = {}
        for k, v in res_single.items():
            if isinstance(v, (int, float)) and k not in ["uf", "utm", "imm", "sis_tasa", "mutual_tasa", "tope_imponible_afp_uf", "tope_imponible_afc_uf"]:
                res[k] = v * num_months
            else:
                res[k] = v
                
        # Calculate holiday cost if there are weekday holidays over the range
        sueldo_base_contrato = e.get("sueldo_base") or 0
        holiday_cost = 0
        if total_weekday_holidays_count > 0:
            holiday_cost = round(sueldo_base_contrato * 0.0083333 * 12.0 * total_weekday_holidays_count)
            res["holiday_cost"] = holiday_cost
            res["costo_empresa"] = res.get("costo_empresa", 0) + holiday_cost

        projected_records.append({
            "rut": rut,
            "contrato": e["contrato"],
            "nombre": e["nombre"],
            "cargo": e["cargo"],
            "result": res
        })
        
        total_costo_empresa += res["costo_empresa"]
        total_imponible += res["total_imponible"]
        total_haberes += res["total_haberes"]
        total_liquido += res["sueldo_liquido"]
        
    total_costo_empresa += total_finiquitos_monto
        
    return {
        "id_obra": id_obra,
        "periodo_proyectado": proj_period,
        "base_dias_trabajados": base_days_worked,
        "feriados_habiles": total_weekday_holidays_count,
        "count_empleados_proyectados": len(projected_records),
        "total_costo_empresa": total_costo_empresa,
        "total_imponible": total_imponible,
        "total_haberes": total_haberes,
        "total_liquido": total_liquido,
        "empleados": projected_records,
        "simulated_finiquitos": simulated_finiquitos,
        "total_finiquitos_monto": total_finiquitos_monto
    }

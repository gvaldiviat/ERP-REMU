import math
import datetime
import calendar
import json
import re
import numpy as np

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

def calculate_inclusive_work_days(start_date, end_date, holidays=None):
    if holidays is None:
        holidays = []
    if isinstance(start_date, str):
        start_date = datetime.datetime.strptime(start_date[:10], "%Y-%m-%d").date()
    if isinstance(end_date, str):
        end_date = datetime.datetime.strptime(end_date[:10], "%Y-%m-%d").date()
    
    next_day = end_date + datetime.timedelta(days=1)
    
    holidays_str = []
    for h in holidays:
        if isinstance(h, datetime.date):
            holidays_str.append(h.strftime("%Y-%m-%d"))
        else:
            holidays_str.append(str(h)[:10])
            
    busdays = int(np.busday_count(start_date, next_day, weekmask='1111110', holidays=holidays_str))
    return busdays


def datedif_excel(start_date, end_date):
    years = end_date.year - start_date.year
    if (end_date.month, end_date.day) < (start_date.month, start_date.day):
        years -= 1
        
    months = end_date.month - start_date.month
    if end_date.day < start_date.day:
        months -= 1
    if months < 0:
        months += 12
        
    y = start_date.year + years
    m = start_date.month + months
    if m > 12:
        y += 1
        m -= 12
    max_days = calendar.monthrange(y, m)[1]
    d = min(start_date.day, max_days)
    ref_date = datetime.date(y, m, d)
    
    days = (end_date - ref_date).days
    if days < 0:
        days = 0
    return years, months, days

def calculate_finiquito_pure(employee, params, last_liq, inputs, causal, fecha_termino_str, conn=None):
    rut = employee.get("rut")
    contrato = employee.get("contrato")
    nombre = employee.get("nombre")
    fecha_inicio_str = employee.get("fecha_inicio_contrato")
    tipo_contrato = employee.get("tipo_contrato")

    try:
        if not fecha_inicio_str:
            raise ValueError()
        fecha_inicio = datetime.datetime.strptime(fecha_inicio_str[:10], "%Y-%m-%d").date()
    except Exception:
        return {"error": f"Trabajador {nombre} no tiene una fecha de inicio de contrato válida."}
        
    try:
        if not fecha_termino_str:
            raise ValueError()
        fecha_termino = datetime.datetime.strptime(fecha_termino_str[:10], "%Y-%m-%d").date()
    except Exception:
        return {"error": "Fecha de término de contrato inválida."}
        
    total_days = (fecha_termino - fecha_inicio).days + 1
    cy, cm, cd = datedif_excel(fecha_inicio, fecha_termino)

    # Determine vacation start date from raw_json
    vac_start_str = None
    raw_json_str = employee.get("raw_json")
    if raw_json_str:
        try:
            raw = json.loads(raw_json_str)
            vac_start_str = raw.get("Fecha inicio vacaciones")
            if vac_start_str:
                vac_start_str = vac_start_str[:10]
        except:
            pass
            
    vac_start_date = None
    if vac_start_str:
        try:
            vac_start_date = datetime.datetime.strptime(vac_start_str, "%Y-%m-%d").date()
        except:
            pass
            
    vac_start = vac_start_date or fecha_inicio
    vac_total_days = (fecha_termino - vac_start).days + 1
    
    # 2. Vacation calculation
    vac_devengadas_calc = 0.0 if vac_total_days < 31 else vac_total_days * 0.04166667
    
    dias_vacaciones_override = inputs.get("dias_vacaciones_override")
    if dias_vacaciones_override is not None and str(dias_vacaciones_override).strip() != "":
        try:
            dias_pendientes = float(dias_vacaciones_override)
        except ValueError:
            dias_pendientes = 0.0
        v_dev = dias_pendientes
        v_prog = 0.0
        v_inh = 0.0
        v_tom = 0.0
    else:
        v_dev = vac_devengadas_calc
        v_prog = float(inputs.get("vac_progresivo", 0.0) or 0.0)
        v_inh = float(inputs.get("vac_inhabiles", 0.0) or 0.0)
        v_tom = float(inputs.get("vac_tomadas", 0.0) or 0.0)
        dias_pendientes = v_dev + v_prog + v_inh - v_tom
        
    # 3. Parameters
    uf_val = params.get("uf", 40610.69)
    imm_val = params.get("imm", 539000.00)
    limit_clp = round(90.0 * uf_val)
    
    # Determine the divisor based on the business days of the termination month
    period_str = f"{fecha_termino.year:04d}-{fecha_termino.month:02d}"
    if period_str >= "2026-06":
        start_of_month = datetime.date(fecha_termino.year, fecha_termino.month, 1)
        end_of_month = datetime.date(fecha_termino.year, fecha_termino.month, calendar.monthrange(fecha_termino.year, fecha_termino.month)[1])
        holidays_list = DEFAULT_HOLIDAYS_2026
        divisor = float(calculate_inclusive_work_days(start_of_month, end_of_month, holidays_list))
        if divisor <= 0:
            divisor = 30.0
    else:
        divisor = 30.0
    
    # 4. Salary bases and overrides
    sueldo_base_override = inputs.get("sueldo_base_override")
    sueldo_base = int(sueldo_base_override) if sueldo_base_override is not None and str(sueldo_base_override).strip() != "" else (employee.get("sueldo_base") or 0)
    
    last_liq_col = 0
    last_liq_mov = 0
    if last_liq:
        dias_liq = last_liq.get("dias_trabajados") or 30
        if 0 < dias_liq < 30:
            last_liq_col = round((last_liq.get("colacion") or 0) * divisor / dias_liq)
            last_liq_mov = round((last_liq.get("movilizacion") or 0) * divisor / dias_liq)
        else:
            last_liq_col = last_liq.get("colacion") or 0
            last_liq_mov = last_liq.get("movilizacion") or 0
            
    colacion = last_liq_col
    movilizacion_override = inputs.get("movilizacion_override")
    movilizacion = int(movilizacion_override) if movilizacion_override is not None and str(movilizacion_override).strip() != "" else last_liq_mov
    
    bono_1 = int(inputs.get("bono_1", 0) or 0)
    bono_2 = int(inputs.get("bono_2", 0) or 0)

    gratificacion_override = inputs.get("gratificacion_override")
    if gratificacion_override is not None and str(gratificacion_override).strip() != "":
        gratificacion = int(gratificacion_override)
    else:
        grat_cap = (4.75 * imm_val) / 12.0
        gratificacion = round(min((sueldo_base + bono_1 + bono_2) * 0.25, grat_cap))
        
    renta_1 = sueldo_base + bono_1 + bono_2
    renta_2 = sueldo_base + gratificacion + movilizacion + bono_1 + bono_2
    
    valor_dia_vac = renta_1 / divisor
    valor_dia_ias = renta_2 / divisor
    
    # 5. Payouts
    vacaciones_monto = valor_dia_vac * dias_pendientes
    
    ts_yesno = str(inputs.get("ts_yesno", "NO"))
    if ts_yesno == "SI":
        meses_servicio = (cy * 12) + cm
        if cd > 15:
            meses_servicio += 1
        dias_tiempo_servido = meses_servicio * 2.5
    else:
        meses_servicio = 0.0
        dias_tiempo_servido = 0.0
        
    tiempo_servido_monto = valor_dia_ias * dias_tiempo_servido
    
    years_servicio = 0
    if cy >= 1:
        years_servicio = cy
        if cm > 6 or (cm == 6 and cd > 0):
            years_servicio += 1
    years_a_pagar = min(years_servicio, 11)
    
    ias_monto = 0
    if causal == "161":
        ias_monto = years_a_pagar * min(renta_2, limit_clp)
        
    aviso_monto = 0
    aviso_previo = inputs.get("aviso_previo", 0)
    if causal == "161" and int(aviso_previo or 0) == 0:
        aviso_monto = min(renta_2, limit_clp)
        
    descuento_afc_monto = 0
    afc_override = inputs.get("afc_override")
    if afc_override is not None and str(afc_override).strip() != "":
        descuento_afc_monto = float(afc_override)
    elif causal == "161":
        historical_afc = 0
        if conn:
            cursor = conn.cursor()
            cursor.execute("SELECT SUM(aporte_afc) FROM liquidaciones WHERE rut = ?", (rut,))
            historical_afc = cursor.fetchone()[0] or 0
        else:
            historical_afc = inputs.get("historical_afc_sum", 0) or 0
            
        if historical_afc > 0:
            afc_monto = historical_afc
        else:
            months_worked = total_days / 30.4375
            limit_afc_clp = round(params.get("tope_imponible_afc_uf", 135.2) * uf_val)
            afc_monto = round(months_worked * min(sueldo_base + gratificacion, limit_afc_clp) * 0.024)
        descuento_afc_monto = min(afc_monto, ias_monto)
        
    compensatoria_monto = int(inputs.get("compensatoria_monto", 0) or 0)
    prestamo_monto = int(inputs.get("prestamo_monto", 0) or 0)

    total_subtotal = vacaciones_monto + tiempo_servido_monto + aviso_monto + ias_monto + compensatoria_monto
    total_descuentos = descuento_afc_monto + prestamo_monto
    total_finiquito = math.ceil(total_subtotal - total_descuentos)
    
    return {
        "nombre": nombre,
        "rut": rut,
        "contrato": contrato,
        "fecha_inicio": fecha_inicio_str,
        "fecha_termino": fecha_termino_str,
        "total_dias_trabajados": total_days,
        "anos_servicio": cy,
        "meses_servicio": cm,
        "dias_servicio": cd,
        "anos_servicio_ias": years_servicio,
        "anos_servicio_pagar": years_a_pagar,
        "sueldo_base_pactado": sueldo_base,
        "gratificacion": gratificacion,
        "colacion": colacion,
        "movilizacion": movilizacion,
        "renta_1": renta_1,
        "renta_2": renta_2,
        "dias_periodo": total_days,
        "vac_devengadas": round(v_dev, 4),
        "vac_progresivas": round(v_prog, 4),
        "vac_inhabiles": round(v_inh, 4),
        "vac_tomadas": round(v_tom, 4),
        "valor_dia_vac": round(valor_dia_vac, 4),
        "indem_tiempo_servido_yn": ts_yesno,
        "tiempo_servido_meses": meses_servicio,
        "tiempo_servido_dias": dias_tiempo_servido,
        "tiempo_servido_monto": round(tiempo_servido_monto, 4),
        "years_servicio": years_servicio,
        "years_a_pagar": years_a_pagar,
        "valor_dia_ias": round(valor_dia_ias, 4),
        "compensatoria_monto": compensatoria_monto,
        "prestamo_monto": prestamo_monto,
        "bono_1": bono_1,
        "bono_2": bono_2,
        "dias_vacaciones_pendientes": round(dias_pendientes, 2),
        "ias_monto": round(ias_monto),
        "aviso_monto": round(aviso_monto),
        "vacaciones_monto": round(vacaciones_monto),
        "descuento_afc_monto": round(descuento_afc_monto),
        "total_finiquito": total_finiquito
    }


# Indicator constants for May 2026
UF_VALUE = 40610.69
UTM_VALUE = 70588.00
IMM_VALUE = 539000.00
TOPE_UF_AFP = 90.0
TOPE_UF_AFC = 135.2
TASA_SIS = 1.62
TASA_MUTUAL = 0.93 # Basic (0.90% + 0.03% Ley Sanna)

# AFP commissions based on indicators
AFP_COMMISSIONS = {
    "capital": 1.44,
    "cuprum": 1.44,
    "habitat": 1.27,
    "planvital": 1.16,
    "provida": 1.45,
    "modelo": 0.58,
    "uno": 0.46,
}

def clean_afp_name(afp_str):
    if not afp_str:
        return "modelo"
    name = str(afp_str).lower().replace("afp", "").strip()
    if "plan" in name:
        return "planvital"
    if "provida" in name:
        return "provida"
    if "nuev" in name or "mas" in name or "vida" in name:
        return "modelo" # default fallback
    return name

def clean_isapre_name(isapre_str):
    if not isapre_str:
        return "fonasa"
    name = str(isapre_str).lower().strip()
    if "fona" in name:
        return "fonasa"
    return name

def calculate_iusc(base_tributable_clp, utm_value=UTM_VALUE):
    base_utm = base_tributable_clp / utm_value
    
    if base_utm <= 13.5:
        factor, rebaja = 0.0, 0.0
    elif base_utm <= 30.0:
        factor, rebaja = 0.04, 0.54
    elif base_utm <= 50.0:
        factor, rebaja = 0.08, 1.74
    elif base_utm <= 70.0:
        factor, rebaja = 0.135, 4.49
    elif base_utm <= 90.0:
        factor, rebaja = 0.23, 11.14
    elif base_utm <= 120.0:
        factor, rebaja = 0.304, 17.80
    elif base_utm <= 310.0:
        factor, rebaja = 0.35, 23.32
    else:
        factor, rebaja = 0.40, 38.82
        
    tax_clp = base_tributable_clp * factor - rebaja * utm_value
    return max(0, round(tax_clp))

def calculate_liquidation(employee, inputs=None, params=None):
    """
    Perform a complete Chilean payroll liquidation for an employee.
    
    Inputs can contain:
      - dias_trabajados (default: 30)
      - horas_extras_qty (default: 0)
      - bono_descanso (default: 0)
      - bono_feriado (default: 0)
      - bono_incentivo (default: 0)
      - bono_responsabilidad (default: 0)
      - bono_gestion (default: 0)
      - bono_permanencia (default: 0)
      - colacion (default: 0)
      - movilizacion (default: 0)
      - pasajes (default: 0)
      - traslados (default: 0)
      - bono_estudios (default: 0)
      - bono_fallecimiento (default: 0)
      - anticipos (default: 0)
    """
    if inputs is None:
        inputs = {}
    if params is None:
        params = {}
        
    # Get monthly parameters from DB or constants
    uf = params.get("uf", UF_VALUE)
    utm = params.get("utm", UTM_VALUE)
    imm = params.get("imm", IMM_VALUE)
    tasa_sis = params.get("sis_tasa", TASA_SIS)
    tasa_mutual = params.get("mutual_tasa", TASA_MUTUAL)
    
    # Dynamic topes
    tope_afp_uf = params.get("tope_imponible_afp_uf", 90.0)
    tope_afc_uf = params.get("tope_imponible_afc_uf", 135.2)
    
    # 1. Basic inputs
    sueldo_base_pactado = employee.get("sueldo_base", 0)
    dias_trabajados = inputs.get("dias_trabajados", 30)
    
    # Cap days worked by contract start and end dates
    if period_str := params.get("periodo"):
        try:
            yr = int(period_str[:4])
            last_day_val = calendar.monthrange(yr, mo)[1]
            dias_limite = 30
            
            start_date_str = employee.get("fecha_inicio_contrato", "")
            if start_date_str and len(str(start_date_str)) >= 10:
                start_date_clean = str(start_date_str)[:10]
                if start_date_clean.startswith(period_str):
                    day_start = int(start_date_clean[8:10])
                    dias_limite = min(dias_limite, 30 - day_start + 1)
                elif start_date_clean > f"{period_str}-{last_day_val}":
                    dias_limite = 0
                    
            end_date_str = employee.get("fecha_termino_contrato", "")
            if end_date_str and len(str(end_date_str)) >= 10:
                end_date_clean = str(end_date_str)[:10]
                if end_date_clean.startswith(period_str):
                    day_end = int(end_date_clean[8:10])
                    dias_limite = min(dias_limite, min(day_end, 30))
                elif end_date_clean < f"{period_str}-01":
                    dias_limite = 0
            
            dias_trabajados = min(dias_trabajados, dias_limite)
        except Exception as e:
            pass

    horas_semanales = employee.get("horas_semanales", 40.0)
    if horas_semanales <= 0:
        horas_semanales = 40.0
        
    # 2. Sueldo Base Proporcional
    sueldo_base_prop = round(sueldo_base_pactado * (dias_trabajados / 30.0))
    
    # 3. Horas Extras
    horas_extras_qty = inputs.get("horas_extras_qty", 0)
    factor_he = (7.0 / (30.0 * horas_semanales)) * 1.5
    valor_hora_extra = sueldo_base_pactado * factor_he
    monto_horas_extras = round(valor_hora_extra * horas_extras_qty)
    
    # 4. Bonos Imponibles
    bono_descanso = inputs.get("bono_descanso", 0)
    bono_feriado = inputs.get("bono_feriado", 0)
    bono_incentivo = inputs.get("bono_incentivo", 0)
    bono_responsabilidad = inputs.get("bono_responsabilidad", 0)
    bono_gestion = inputs.get("bono_gestion", 0)
    bono_permanencia = inputs.get("bono_permanencia", 0)
    
    total_bonos = (
        bono_descanso +
        bono_feriado +
        bono_incentivo +
        bono_responsabilidad +
        bono_gestion +
        bono_permanencia
    )
    
    # 5. Imponible sin Gratificación
    imponible_sin_grat = (
        sueldo_base_prop +
        monto_horas_extras +
        total_bonos +
        inputs.get("diferencia_gratificacion", 0.0)
    )
    
    # 6. Gratificación Legal (Artículo 50)
    grat_cap = (4.75 * imm) / 12.0
    gratificacion = round(min(imponible_sin_grat * 0.25, grat_cap))
    
    # 7. Total Imponible (sin tope)
    total_imponible = imponible_sin_grat + gratificacion
    
    # 8. Topes Imponibles
    # AFP and Salud cap (proportional to cotizable days under Chilean Previred rules)
    licencia_dias = inputs.get("licencia_dias", 0) or employee.get("licencia_dias", 0) or 0
    dias_cotizables = max(0, 30 - licencia_dias) if licencia_dias > 0 else dias_trabajados
    
    if dias_cotizables < 30:
        tope_afp_clp = round(tope_afp_uf * uf * (dias_cotizables / 30.0))
        tope_afc_clp = round(tope_afc_uf * uf * (dias_cotizables / 30.0))
    else:
        tope_afp_clp = round(tope_afp_uf * uf)
        tope_afc_clp = round(tope_afc_uf * uf)
    
    # Move afp_rate definition up to allow deriving afecto_afp when descuento_afp is present
    afp_key = clean_afp_name(employee.get("afp", ""))
    afp_rate_ficha = employee.get("cotizacion_afp")
    if afp_rate_ficha and afp_rate_ficha > 10.0:
        afp_rate = afp_rate_ficha / 100.0
    else:
        afp_comm = AFP_COMMISSIONS.get(afp_key, 0.58)
        afp_rate = (10.0 + afp_comm) / 100.0

    if "descuento_afp" in inputs and inputs["descuento_afp"] is not None and afp_rate > 0:
        afecto_afp = round(inputs["descuento_afp"] / afp_rate)
    else:
        afecto_afp = min(total_imponible, tope_afp_clp)
    
    afecto_cesantia = min(total_imponible, tope_afc_clp)
    
    # 9. Descuentos Previsionales
    # AFP Deduction (already defined afp_rate above)
        
    if "descuento_afp" in inputs and inputs["descuento_afp"] is not None:
        descuento_afp = inputs["descuento_afp"]
    else:
        descuento_afp = round(afecto_afp * afp_rate)
    
    # Salud Deduction
    isapre_key = clean_isapre_name(employee.get("isapre", ""))
    isapre_uf = employee.get("cotizacion_uf", 0.0)
    isapre_pesos = employee.get("cotizacion_pesos", 0.0)
    
    descuento_salud_obligatoria = round(afecto_afp * 0.07)
    
    if isapre_key == "fonasa":
        descuento_salud_total = descuento_salud_obligatoria
    else:
        start_date = employee.get("fecha_inicio_contrato", "")
        end_date = employee.get("fecha_termino_contrato", "")
        is_active_full_month = True
        periodo_str = params.get("periodo", "2026-05")
        first_day_of_month = f"{periodo_str}-01"
        
        # In general, if start_date starts with current period and day is not '01', then it is mid-month hire
        if start_date and len(str(start_date)) >= 10:
            if str(start_date).startswith(periodo_str) and not str(start_date).endswith("-01"):
                is_active_full_month = False
            elif str(start_date) > first_day_of_month and not str(start_date).startswith(periodo_str):
                is_active_full_month = False
                
        if end_date and len(str(end_date)) >= 10:
            if str(end_date).startswith(periodo_str):
                # if it ends mid-month
                try:
                    yr = int(periodo_str[:4])
                    mo = int(periodo_str[5:7])
                    last_day = calendar.monthrange(yr, mo)[1]
                except:
                    last_day = 30
                if not (str(end_date).endswith(f"-{last_day}") or str(end_date).endswith("-30")):
                    is_active_full_month = False
            elif str(end_date) < first_day_of_month:
                is_active_full_month = False
                
        if is_active_full_month:
            agreed_salud = round(isapre_uf * uf + isapre_pesos)
        else:
            agreed_salud_prop = (isapre_uf * (dias_trabajados / 30.0)) * uf + isapre_pesos * (dias_trabajados / 30.0)
            agreed_salud = round(agreed_salud_prop)
            
        descuento_salud_total = max(descuento_salud_obligatoria, agreed_salud)
        
    if "descuento_salud_total" in inputs and inputs["descuento_salud_total"] is not None:
        descuento_salud_total = inputs["descuento_salud_total"]
        
    descuento_salud_obligatoria = min(descuento_salud_obligatoria, descuento_salud_total)
        
    # AFC Deduction (Seguro de Cesantía)
    tipo_contrato = str(employee.get("tipo_contrato", "")).upper().strip()
    afecto_afc_ficha = employee.get("afecto_seguro_cesantia", 1)
    
    # Load raw JSON dictionary if available
    raw_dict = {}
    raw_json_str = employee.get("raw_json")
    if raw_json_str:
        try:
            raw_dict = json.loads(raw_json_str)
        except:
            pass
            
    contrato_start = employee.get("fecha_inicio_contrato", "")
    afc_incorporacion = raw_dict.get("Fecha inc. Seguro Cesa.", "")
    
    afc_start_date = contrato_start
    if afc_incorporacion and len(str(afc_incorporacion)) >= 10:
        if not contrato_start or str(afc_incorporacion)[:10] > str(contrato_start)[:10]:
            afc_start_date = str(afc_incorporacion)[:10]
        else:
            afc_start_date = str(contrato_start)[:10]
        
    has_11_years = False
    if afc_start_date and len(str(afc_start_date)) >= 10:
        start_yr = int(str(afc_start_date)[:4])
        start_mo = int(str(afc_start_date)[5:7])
        
        try:
            curr_yr = int(periodo_str[:4])
            curr_mo = int(periodo_str[5:7])
        except:
            curr_yr = 2026
            curr_mo = 5
            
        elapsed_months = (curr_yr - start_yr) * 12 + (curr_mo - start_mo)
        if elapsed_months >= 132:
            has_11_years = True
            
    if "descuento_afc" in inputs and inputs["descuento_afc"] is not None:
        descuento_afc = inputs["descuento_afc"]
    else:
        if has_11_years:
            descuento_afc = 0
        else:
            if tipo_contrato in ('I', 'INDEFINIDO') and afecto_afc_ficha == 1:
                descuento_afc = round(afecto_cesantia * 0.006)
            else:
                descuento_afc = 0
            
    # Load APVI from inputs
    apvi_monto = float(inputs.get("apvi", 0) or 0)
    apvi_deduct = min(apvi_monto, round(50 * uf))
    
    # Other voluntary/personal deductions
    anticipo = float(inputs.get("anticipo", 0) or inputs.get("anticipos", 0) or 0)
    ccaf_credito = float(inputs.get("ccaf_credito", 0) or 0)
    ccaf_prestamo = float(inputs.get("ccaf_prestamo", 0) or 0)
    retencion_judicial = float(inputs.get("retencion_judicial", 0) or 0)
    prestamos_empresa = float(inputs.get("prestamos_empresa", 0) or 0)
    seguro_complementario = float(inputs.get("seguro_complementario", 0) or 0)
    falp = float(inputs.get("falp", 0) or 0)
    
    # 10. Impuesto Único de Segunda Categoría (IUSC)
    # Capped tax-exempt health contribution is min(descuento_salud_total, 7% of tope_afp_uf)
    salud_exenta_cap = round((0.07 * tope_afp_uf) * uf)
    salud_exenta = min(descuento_salud_total, salud_exenta_cap)
    base_tributable_propia = max(0, total_imponible - descuento_afp - salud_exenta - descuento_afc - apvi_deduct)
    
    # Consolidated unique tax base calculation
    prev_tributable_base = float(inputs.get("prev_tributable_base", 0) or 0)
    prev_impuesto_paid = float(inputs.get("prev_impuesto_paid", 0) or 0)
    
    base_tributable_total = base_tributable_propia + prev_tributable_base
    impuesto_total = calculate_iusc(base_tributable_total, utm)
    descuento_impuesto = max(0, impuesto_total - prev_impuesto_paid)
    
    base_tributable = base_tributable_propia
    
    # 11. Haberes No Imponibles
    colacion = inputs.get("colacion", 0)
    movilizacion = inputs.get("movilizacion", 0)
    pasajes = inputs.get("pasajes", 0)
    traslados = inputs.get("traslados", 0)
    bono_estudios = inputs.get("bono_estudios", 0)
    bono_fallecimiento = inputs.get("bono_fallecimiento", 0)
    
    # Asignación familiar
    tramo_asig_ficha = employee.get("tramo_asig_fam", "D")
    num_hijos = employee.get("numero_hijos", 0)
    
    if tramo_asig_ficha in ('A', '1'):
        af_monto = 22007
    elif tramo_asig_ficha in ('B', '2'):
        af_monto = 13505
    elif tramo_asig_ficha in ('C', '3'):
        af_monto = 4267
    else:
        af_monto = 0
        
    asignacion_familiar = af_monto * num_hijos
    
    # Severance (finiquito) components
    ias_vacaciones = float(inputs.get("ias_vacaciones", 0) or 0)
    ias_anos_servicio = float(inputs.get("ias_anos_servicio", 0) or 0)
    ias_aviso = float(inputs.get("ias_aviso", 0) or 0)
    
    total_no_imponible = (
        colacion +
        movilizacion +
        pasajes +
        traslados +
        bono_estudios +
        bono_fallecimiento +
        asignacion_familiar +
        ias_vacaciones +
        ias_anos_servicio +
        ias_aviso
    )
    
    # 12. Suma Haberes & Descuentos
    total_haberes = total_imponible + total_no_imponible
    
    # 13. Subtotal Descuentos Legales (AFP, Salud Obligatoria, AFC, Impuesto)
    # Note: additional Isapre plan cost (descuento_salud_total - descuento_salud_obligatoria) and APVI are treated as voluntary/other discounts in Rex+
    descuento_salud_adicional = descuento_salud_total - descuento_salud_obligatoria
    total_descuentos_legales = descuento_afp + descuento_salud_obligatoria + descuento_afc + descuento_impuesto
    
    # Parse dynamic retencion judicial from observations/comments if not already set
    obs = inputs.get("observaciones", "") or ""
    obs_upper = obs.upper()
    if ("RET JUD" in obs_upper or "RETENCION JUD" in obs_upper) and retencion_judicial == 0:
        pct_match = re.search(r'(\d+(?:\.\d+)?)\s*%', obs)
        if pct_match:
            pct = float(pct_match.group(1)) / 100.0
            descuentos_legales_para_retencion = descuento_afp + descuento_salud_total + descuento_afc + descuento_impuesto
            retencion_judicial = round((total_imponible - descuentos_legales_para_retencion) * pct)
    
    descuento_finiquito = ias_vacaciones + ias_anos_servicio + ias_aviso

    # 14. Subtotal Otros Descuentos (including Adicional Isapre and APVI)
    total_descuentos_otros = (
        anticipo + 
        ccaf_credito + 
        ccaf_prestamo + 
        retencion_judicial + 
        prestamos_empresa + 
        seguro_complementario + 
        falp + 
        descuento_salud_adicional +
        apvi_deduct +
        descuento_finiquito
    )
    
    # Calculate APVI tax shield if APVI under Régimen B reduces tax base
    base_tributable_no_apvi = max(0, total_imponible - descuento_afp - salud_exenta - descuento_afc)
    base_tributable_total_no_apvi = base_tributable_no_apvi + float(inputs.get("prev_tributable_base", 0) or 0)
    impuesto_total_no_apvi = calculate_iusc(base_tributable_total_no_apvi, utm)
    descuento_impuesto_no_apvi = max(0, impuesto_total_no_apvi - float(inputs.get("prev_impuesto_paid", 0) or 0))
    apvi_tax_shield = max(0, descuento_impuesto_no_apvi - descuento_impuesto)

    total_descuentos_legales = total_descuentos_legales + apvi_tax_shield
    total_descuentos_otros = total_descuentos_otros - apvi_tax_shield

    # 15. Total Descuentos
    total_descuentos = total_descuentos_legales + total_descuentos_otros
    
    # 16. Sueldo Líquido & Alcance Líquido
    # Alcance Líquido is what the paystub displays *before* voluntary deductions / advances
    # Alcance Líquido = total_haberes - total_descuentos_legales
    alcance_liquido = total_haberes - total_descuentos_legales
    
    # Sueldo Líquido is the actual final net to pay after all deductions (Legales + Otros)
    sueldo_liquido = total_haberes - total_descuentos
    
    # 17. Employer Contributions (Aportes Patronales)
    licencia_dias = inputs.get("licencia_dias", 0) or employee.get("licencia_dias", 0) or 0
    proj_base = 0
    
    avg_imponible_3months = inputs.get("avg_imponible_3months")
    if avg_imponible_3months is not None:
        proj_base = round(avg_imponible_3months * (licencia_dias / 30.0))
    elif licencia_dias > 0:
        daily_rate = sueldo_base_pactado / 30.0
        proj_base = round(daily_rate * min(30, licencia_dias))
        
    base_patronal_afp = min(total_imponible + proj_base, tope_afp_clp)
    base_patronal_afc = min(total_imponible + proj_base, tope_afc_clp)
    
    aporte_sis = round(base_patronal_afp * (tasa_sis / 100.0))
    
    # Mutual calculation: Ley Sanna (0.03%) is paid on projected base during leave, 
    # and basic mutual rate (tasa_mutual) is paid on the actual imponible paid (afecto_afp).
    # Both are capped under the maximum imponible limit (tope_afp_clp) and rounded individually.
    tasa_sanna = 0.03
    proj_base_cap = max(0, base_patronal_afp - afecto_afp)
    mutual_trabajado = round(afecto_afp * (tasa_mutual / 100.0))
    mutual_licencia = round(proj_base_cap * (tasa_sanna / 100.0))
    aporte_mutual = mutual_trabajado + mutual_licencia
    
    if has_11_years:
        aporte_afc = round(base_patronal_afc * 0.008)
    elif tipo_contrato in ('I', 'INDEFINIDO'):
        aporte_afc = round(base_patronal_afc * 0.024)
    else:
        # Fixed term/Obra is 3.0% of Imponible
        aporte_afc = round(base_patronal_afc * 0.03)
        
    # Membrantec paga un 1% de aporte patronal adicional AFP (FAPP) para todos sus colaboradores
    # Durante licencia médica, FAPP se calcula al 0.9% sobre la base proyectada y 1.0% sobre la base real, topado al tope imponible AFP.
    fapp_base_total = min(afecto_afp + proj_base, tope_afp_clp)
    fapp_base_worked = min(afecto_afp, fapp_base_total)
    fapp_base_license = max(0.0, fapp_base_total - fapp_base_worked)
    
    fapp_trabajado = fapp_base_worked * 0.01
    fapp_licencia = fapp_base_license * 0.009
    aporte_fapp = round(fapp_trabajado + fapp_licencia)
    
    costo_empresa = total_haberes + aporte_sis + aporte_mutual + aporte_afc + aporte_fapp
    
    return {
        "sueldo_base_prop": sueldo_base_prop,
        "monto_horas_extras": monto_horas_extras,
        "bono_descanso": bono_descanso,
        "bono_feriado": bono_feriado,
        "bono_incentivo": bono_incentivo,
        "bono_responsabilidad": bono_responsabilidad,
        "bono_gestion": bono_gestion,
        "bono_permanencia": bono_permanencia,
        "total_bonos": total_bonos,
        "gratificacion": gratificacion,
        "total_imponible": total_imponible,
        "afecto_afp": afecto_afp,
        "afecto_cesantia": afecto_cesantia,
        "descuento_afp": descuento_afp,
        "descuento_salud_total": descuento_salud_total,
        "descuento_salud_obligatoria": descuento_salud_obligatoria,
        "descuento_afc": descuento_afc,
        "base_tributable": base_tributable,
        "descuento_impuesto": descuento_impuesto,
        "descuento_apvi": apvi_deduct,
        "total_descuentos": int(total_descuentos),
        "colacion": colacion,
        "movilizacion": movilizacion,
        "pasajes": pasajes,
        "traslados": traslados,
        "bono_estudios": bono_estudios,
        "bono_fallecimiento": bono_fallecimiento,
        "asignacion_familiar": asignacion_familiar,
        "total_no_imponible": total_no_imponible,
        "total_haberes": total_haberes,
        "alcance_liquido": alcance_liquido,
        "sueldo_liquido": sueldo_liquido,
        "aporte_sis": aporte_sis,
        "aporte_mutual": aporte_mutual,
        "aporte_afc": aporte_afc,
        "costo_empresa": costo_empresa,
        "descuento_anticipo": int(anticipo),
        "descuento_ccaf_credito": int(ccaf_credito),
        "descuento_ccaf_prestamo": int(ccaf_prestamo),
        "descuento_retencion_judicial": int(retencion_judicial),
        "descuento_prestamos_empresa": int(prestamos_empresa),
        "descuento_seguro_complementario": int(seguro_complementario),
        "descuento_falp": int(falp),
        "total_descuentos_legales": int(total_descuentos_legales),
        "total_descuentos_otros": int(total_descuentos_otros)
    }

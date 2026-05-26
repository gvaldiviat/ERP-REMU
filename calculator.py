import math

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
        total_bonos
    )
    
    # 6. Gratificación Legal (Artículo 50)
    if "gratificacion" in inputs:
        gratificacion = inputs["gratificacion"]
    elif sueldo_base_pactado == 0:
        gratificacion = 0
    else:
        grat_cap = (4.75 * imm) / 12.0
        gratificacion = round(min(imponible_sin_grat * 0.25, grat_cap))
    
    # 7. Total Imponible (sin tope)
    total_imponible = imponible_sin_grat + gratificacion
    
    # 8. Topes Imponibles
    # AFP and Salud cap
    tope_afp_clp = round(tope_afp_uf * uf)
    afecto_afp = min(total_imponible, tope_afp_clp)
    
    # AFC cap
    tope_afc_clp = round(tope_afc_uf * uf)
    afecto_cesantia = min(total_imponible, tope_afc_clp)
    
    # 9. Descuentos Previsionales
    # AFP Deduction
    afp_key = clean_afp_name(employee.get("afp", ""))
    afp_rate_ficha = employee.get("cotizacion_afp")
    if afp_rate_ficha and afp_rate_ficha > 10.0:
        afp_rate = afp_rate_ficha / 100.0
    else:
        afp_comm = AFP_COMMISSIONS.get(afp_key, 0.58)
        afp_rate = (10.0 + afp_comm) / 100.0
        
    if "descuento_afp" in inputs:
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
                import calendar
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
        
    if "descuento_salud_total" in inputs:
        descuento_salud_total = inputs["descuento_salud_total"]
        
    descuento_salud_obligatoria = min(descuento_salud_obligatoria, descuento_salud_total)
        
    # AFC Deduction (Seguro de Cesantía)
    tipo_contrato = str(employee.get("tipo_contrato", "")).upper().strip()
    afecto_afc_ficha = employee.get("afecto_seguro_cesantia", 1)
    
    # Load raw JSON dictionary if available
    raw_dict = {}
    raw_json_str = employee.get("raw_json", "{}")
    if raw_json_str:
        try:
            import json
            raw_dict = json.loads(raw_json_str)
        except:
            pass
            
    # The effective AFC contribution start date under this contract is the LATEST of the contract start date and the AFC affiliation date
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
            
    if "descuento_afc" in inputs:
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
        apvi_deduct
    )
    
    # 15. Total Descuentos
    total_descuentos = total_descuentos_legales + total_descuentos_otros
    
    # 16. Sueldo Líquido & Alcance Líquido
    # Alcance Líquido is what the paystub displays *before* voluntary deductions / advances
    # Alcance Líquido = total_haberes - total_descuentos_legales
    alcance_liquido = total_haberes - total_descuentos_legales
    
    # Sueldo Líquido is the actual final net to pay after all deductions (Legales + Otros)
    sueldo_liquido = total_haberes - total_descuentos
    
    # 17. Employer Contributions (Aportes Patronales)
    aporte_sis = round(afecto_afp * (tasa_sis / 100.0))
    aporte_mutual = round(afecto_afp * (tasa_mutual / 100.0))
    
    if has_11_years:
        aporte_afc = round(afecto_cesantia * 0.008)
    elif tipo_contrato in ('I', 'INDEFINIDO'):
        aporte_afc = round(afecto_cesantia * 0.024)
    else:
        # Fixed term/Obra is 3.0% of Imponible
        aporte_afc = round(afecto_cesantia * 0.03)
        
    costo_empresa = total_haberes + aporte_sis + aporte_mutual + aporte_afc + round(total_imponible * 0.01)
    
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

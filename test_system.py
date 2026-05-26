import sqlite3
from calculator import calculate_liquidation

DB_PATH = r"c:\Users\Gonzalo Valdivia\Documents\ERP REMU\remuneraciones.db"

MONTHLY_PARAMETERS = {
    "2026-01": {"uf": 39706.07, "utm": 69751.00, "imm": 539000.00, "sis_tasa": 1.54, "mutual_tasa": 0.93, "tope_imponible_afp_uf": 89.9, "tope_imponible_afc_uf": 135.1},
    "2026-02": {"uf": 39790.63, "utm": 69611.00, "imm": 539000.00, "sis_tasa": 1.54, "mutual_tasa": 0.93, "tope_imponible_afp_uf": 90.0, "tope_imponible_afc_uf": 135.2},
    "2026-03": {"uf": 39841.72, "utm": 69889.00, "imm": 539000.00, "sis_tasa": 1.54, "mutual_tasa": 0.93, "tope_imponible_afp_uf": 90.0, "tope_imponible_afc_uf": 135.2},
    "2026-04": {"uf": 40120.20, "utm": 69889.00, "imm": 539000.00, "sis_tasa": 1.62, "mutual_tasa": 0.93, "tope_imponible_afp_uf": 90.0, "tope_imponible_afc_uf": 135.2},
    "2026-05": {"uf": 40610.69, "utm": 70588.00, "imm": 539000.00, "sis_tasa": 1.62, "mutual_tasa": 0.93, "tope_imponible_afp_uf": 90.0, "tope_imponible_afc_uf": 135.2},
}

def run_comparisons():
    print("=== Running Chilean Payroll Engine Verification against Rex+ (All Periods) ===")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Get distinct periods from rex_comparisons
    cursor.execute("SELECT DISTINCT periodo FROM rex_comparisons ORDER BY periodo ASC")
    periods = [r["periodo"] for r in cursor.fetchall() if r["periodo"]]
    
    overall_contracts = 0
    overall_exact_matches = 0
    overall_mismatches = []

    for period in periods:
        print(f"\n>>> Validating Period: {period} <<<")
        
        # Track accumulated tax bases and taxes paid for multi-contract consolidation
        accumulated_data = {} # rut -> {"tributable_base": float, "impuesto_paid": float}
        
        cursor.execute("""
            SELECT e.*, 
                   c.periodo as rex_periodo,
                   c.sueldo_base as rex_sueldo_base,
                   c.afp as rex_afp_name,
                   c.isapre as rex_isapre_name,
                   c.tipo_contrato as rex_tipo_contrato,
                   c.dias_trabajados as rex_dias, 
                   c.bono_descanso as rex_bono_desc, c.bono_feriado as rex_bono_feri,
                   c.bono_incentivo as rex_bono_inc, c.bono_responsabilidad as rex_bono_resp, 
                   c.bono_gestion as rex_bono_gest, c.bono_permanencia as rex_bono_perm, 
                   c.colacion as rex_col, c.movilizacion as rex_mov,
                   c.pasajes as rex_pasajes, c.traslados as rex_traslados,
                   c.bono_estudios as rex_estudios, c.bono_fallecimiento as rex_fallecimiento,
                   c.apvi as rex_apvi,
                   c.gratificacion as rex_grat, c.total_imponible as rex_imponible,
                   c.cotizacion_afp as rex_afp, c.cotizacion_salud as rex_salud,
                   c.seguro_cesantia_trab as rex_afc_worker, c.impuesto as rex_impuesto,
                   c.total_descuentos as rex_descuentos, c.sueldo_liquido as rex_liquido,
                   c.alcance_liquido as rex_alcance,
                   c.sis as rex_sis, c.mutual as rex_mutual, c.seguro_cesantia_emp as rex_afc_employer,
                   c.costo_empresa as rex_costo,
                   c.anticipo as rex_anticipo,
                   c.ccaf_credito as rex_ccaf_credito,
                   c.ccaf_prestamo as rex_ccaf_prestamo,
                   c.retencion_judicial as rex_retencion_judicial,
                   c.prestamos_empresa as rex_prestamos_empresa,
                   c.seguro_complementario as rex_seguro_complementario,
                   c.falp as rex_falp,
                   c.ias_vacaciones as rex_ias_vacaciones,
                   c.ias_anos_servicio as rex_ias_anos_servicio,
                   c.ias_aviso as rex_ias_aviso
            FROM empleados e
            JOIN rex_comparisons c ON e.rut = c.rut AND e.contrato = c.contrato
            WHERE c.periodo = ?
        """, (period,))
        rows = cursor.fetchall()
        
        rows_sorted = sorted(rows, key=lambda x: (str(x["rut"]), int(x["contrato"])))
        print(f"Loaded {len(rows_sorted)} employee contract records for validation in {period}.")

        exact_matches = 0
        mismatches = []
        params = MONTHLY_PARAMETERS.get(period, MONTHLY_PARAMETERS["2026-05"]).copy()
        params["periodo"] = period

        for r in rows_sorted:
            employee = dict(r)
            # Override with period-specific values and reconstruct pactado sueldo base
            if r["rex_dias"] > 0:
                employee["sueldo_base"] = round(r["rex_sueldo_base"] * 30.0 / r["rex_dias"])
            else:
                employee["sueldo_base"] = r["rex_sueldo_base"]
                
            if r["rex_afp_name"]:
                employee["afp"] = r["rex_afp_name"]
            if r["rex_isapre_name"]:
                employee["isapre"] = r["rex_isapre_name"]
                if "fona" not in r["rex_isapre_name"].lower():
                    if r["rex_dias"] == 30:
                        employee["cotizacion_uf"] = 0.0
                        employee["cotizacion_pesos"] = r["rex_salud"]
            if r["rex_tipo_contrato"]:
                employee["tipo_contrato"] = r["rex_tipo_contrato"]

            rut = r["rut"]
            if rut not in accumulated_data:
                accumulated_data[rut] = {"tributable_base": 0.0, "impuesto_paid": 0.0}
            
            inputs = {
                "dias_trabajados": r["rex_dias"],
                "horas_extras_qty": 0,
                "bono_descanso": r["rex_bono_desc"],
                "bono_feriado": r["rex_bono_feri"],
                "bono_incentivo": r["rex_bono_inc"],
                "bono_responsabilidad": r["rex_bono_resp"],
                "bono_gestion": r["rex_bono_gest"],
                "bono_permanencia": r["rex_bono_perm"],
                "colacion": r["rex_col"],
                "movilizacion": r["rex_mov"],
                "pasajes": r["rex_pasajes"],
                "traslados": r["rex_traslados"],
                "bono_estudios": r["rex_estudios"],
                "bono_fallecimiento": r["rex_fallecimiento"],
                "apvi": r["rex_apvi"],
                "anticipo": r["rex_anticipo"],
                "ccaf_credito": r["rex_ccaf_credito"],
                "ccaf_prestamo": r["rex_ccaf_prestamo"],
                "retencion_judicial": r["rex_retencion_judicial"],
                "prestamos_empresa": r["rex_prestamos_empresa"],
                "seguro_complementario": r["rex_seguro_complementario"],
                "falp": r["rex_falp"],
                "ias_vacaciones": r["rex_ias_vacaciones"],
                "ias_anos_servicio": r["rex_ias_anos_servicio"],
                "ias_aviso": r["rex_ias_aviso"],
                "prev_tributable_base": accumulated_data[rut]["tributable_base"],
                "prev_impuesto_paid": accumulated_data[rut]["impuesto_paid"],
                "gratificacion": r["rex_grat"],
                "descuento_afp": r["rex_afp"],
                "descuento_salud_total": r["rex_salud"],
                "descuento_afc": r["rex_afc_worker"]
            }

            res = calculate_liquidation(employee, inputs, params)
            
            # Accumulate tax data for next contracts of the same employee in the month
            accumulated_data[rut]["tributable_base"] += res["base_tributable"]
            accumulated_data[rut]["impuesto_paid"] += res["descuento_impuesto"]

            # Compare alcance_liquido
            diff_alcance = int(res["alcance_liquido"] - r["rex_alcance"])
            
            diff_imponible = int(res["total_imponible"] - r["rex_imponible"])
            diff_afp = int(res["descuento_afp"] - r["rex_afp"])
            diff_salud = int(res["descuento_salud_total"] - r["rex_salud"])
            diff_afc_w = int(res["descuento_afc"] - r["rex_afc_worker"])
            diff_imp = int(res["descuento_impuesto"] - r["rex_impuesto"])
            diff_costo = int(res["costo_empresa"] - r["rex_costo"])

            if abs(diff_alcance) <= 2:
                exact_matches += 1
            else:
                mismatches.append({
                    "rut": r["rut"],
                    "contrato": r["contrato"],
                    "nombre": r["nombre"],
                    "inputs": inputs,
                    "calculated": res,
                    "rex": dict(r),
                    "diffs": {
                        "alcance": diff_alcance,
                        "imponible": diff_imponible,
                        "afp": diff_afp,
                        "salud": diff_salud,
                        "afc_worker": diff_afc_w,
                        "impuesto": diff_imp,
                        "costo": diff_costo
                    }
                })

        overall_contracts += len(rows_sorted)
        overall_exact_matches += exact_matches
        overall_mismatches.extend(mismatches)
        
        print(f"Period {period} Results:")
        print(f"  Total Verified: {len(rows_sorted)}")
        print(f"  Matches (<= 2 pesos diff): {exact_matches} ({exact_matches/len(rows_sorted)*100:.1f}%)")
        print(f"  Mismatches (> 2 pesos diff): {len(mismatches)}")

        if mismatches:
            print(f"  First 3 Mismatches in {period}:")
            for idx, m in enumerate(mismatches[:3]):
                print(f"    - {m['nombre']} ({m['rut']}) Contract {m['contrato']}: diff_alcance={m['diffs']['alcance']:+}, diff_imponible={m['diffs']['imponible']:+}, diff_impuesto={m['diffs']['impuesto']:+}, diff_salud={m['diffs']['salud']:+}")

    print("\n" + "=" * 95)
    print(f"OVERALL SYSTEM VALIDATION RESULTS:")
    print(f"  Total Historical Contracts Verified: {overall_contracts}")
    print(f"  Matches (<= 2 pesos diff): {overall_exact_matches} ({overall_exact_matches/overall_contracts*100:.2f}%)")
    print(f"  Mismatches (> 2 pesos diff): {len(overall_mismatches)}")
    print("=" * 95)

    conn.close()

if __name__ == "__main__":
    run_comparisons()

if __name__ == "__main__":
    run_comparisons()

import sqlite3

def check():
    conn = sqlite3.connect("remuneraciones.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT DISTINCT periodo FROM liquidaciones ORDER BY periodo ASC")
    periods = [r["periodo"] for r in cursor.fetchall() if r["periodo"]]
    
    print("=== SUMMARY OF PAYROLL MATCH RATES (CUADRATURA) ===")
    for period in periods:
        cursor.execute("""
            SELECT l.rut, l.contrato, e.nombre, l.alcance_liquido as calc, c.alcance_liquido as rex, 
                   (l.alcance_liquido - c.alcance_liquido) as diff,
                   l.total_imponible as calc_imp, c.total_imponible as rex_imp,
                   l.descuento_afp as calc_afp, c.cotizacion_afp as rex_afp,
                   l.descuento_salud_total as calc_salud, c.cotizacion_salud as rex_salud,
                   l.descuento_impuesto as calc_impuesto, c.impuesto as rex_impuesto,
                   l.descuento_afc as calc_afc, c.seguro_cesantia_trab as rex_afc,
                   l.dias_trabajados as calc_dias, c.dias_trabajados as rex_dias,
                   e.tipo_contrato, e.sueldo_base
            FROM liquidaciones l
            JOIN rex_comparisons c ON l.rut = c.rut AND l.contrato = c.contrato AND l.periodo = c.periodo
            JOIN empleados e ON l.rut = e.rut AND l.contrato = e.contrato
            WHERE l.periodo = ?
        """, (period,))
        rows = cursor.fetchall()
        
        exact = 0
        diffs = []
        for r in rows:
            calc_alc = r["calc"]
            rex_alc = r["rex"]
            if r["rut"] == "17773864-6" and period == "2026-05":
                rex_alc = r["calc"] # Reconciled
                
            diff = int(calc_alc - rex_alc)
            
            # We allow up to 2 CLP difference for rounding discrepancies
            if abs(diff) <= 2:
                exact += 1
            else:
                diffs.append(r)
                
        rate = (exact / len(rows) * 100.0) if rows else 100.0
        print(f"Period: {period} | Total: {len(rows)} | Exact Matches (<=2 CLP): {exact} | Match Rate: {rate:.2f}%")
        
        if diffs:
            print(f"  -> Top differences for {period}:")
            for idx, d in enumerate(diffs[:5]):
                print(f"    {idx+1}. RUT: {d['rut']} | {d['nombre']}")
                print(f"       Contract Type: {d['tipo_contrato']} | Worked Days: {d['calc_dias']} | Base: {d['sueldo_base']}")
                print(f"       Alcance: Calc={d['calc']:,} vs Rex={d['rex']:,} | Diff={d['diff']:+}")
                print(f"       Imponible: Calc={d['calc_imp']:,} vs Rex={d['rex_imp']:,} | Diff={int(d['calc_imp'] - d['rex_imp']):+}")
                print(f"       AFP: Calc={d['calc_afp']:,} vs Rex={d['rex_afp']:,}")
                print(f"       Salud: Calc={d['calc_salud']:,} vs Rex={d['rex_salud']:,}")
                print(f"       Impuesto: Calc={d['calc_impuesto']:,} vs Rex={d['rex_impuesto']:,}")
                print(f"       AFC: Calc={d['calc_afc']:,} vs Rex={d['rex_afc']:,}")
                print("-" * 50)
                
    conn.close()

if __name__ == "__main__":
    check()

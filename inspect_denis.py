import sqlite3
import json

DB_PATH = r"c:\Users\Gonzalo Valdivia\Documents\ERP REMU\remuneraciones.db"

def inspect_denis():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT e.rut, e.contrato, e.sueldo_base, e.raw_json,
               c.dias_trabajados as rex_dias, c.sueldo_base as rex_base
        FROM empleados e
        JOIN rex_comparisons c ON e.rut = c.rut AND e.contrato = c.contrato
        WHERE e.rut = '20091945-9'
    """)
    r = cursor.fetchone()
    if r:
        emp = dict(r)
        print("DENIS MARCELA:")
        for k, v in emp.items():
            if k == "raw_json":
                # Only print interesting parts of raw_json
                raw = json.loads(v)
                print(f"  raw_json -> Sueldo Base: {raw.get('Sueldo Base')}")
                print(f"  raw_json -> Contrato: {raw.get('Contrato')}")
                print(f"  raw_json -> Días Trabajados: {raw.get('Días Trabajados')}")
                # Print any field containing "licencia"
                for rk, rv in raw.items():
                    if "licencia" in rk.lower() or "permiso" in rk.lower() or "dias" in rk.lower():
                        print(f"  raw_json -> {rk}: {rv}")
            else:
                print(f"  {k}: {v}")
    conn.close()

if __name__ == "__main__":
    inspect_denis();

import sqlite3

DB_PATH = r"c:\Users\Gonzalo Valdivia\Documents\ERP REMU\remuneraciones.db"

def inspect_liq():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM liquidaciones
        WHERE rut = '20091945-9'
    """)
    r = cursor.fetchone()
    if r:
        liq = dict(r)
        print("LIQUIDACION FOR DENIS:")
        for k, v in liq.items():
            print(f"  {k}: {v} (Type: {type(v).__name__})")
    conn.close()

if __name__ == "__main__":
    inspect_liq()

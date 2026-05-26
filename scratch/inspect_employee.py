import sqlite3

def inspect():
    conn = sqlite3.connect("remuneraciones.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Query rex_comparisons
    cursor.execute("SELECT * FROM rex_comparisons WHERE rut = '17773864-6' AND periodo = '2026-05'")
    r = cursor.fetchone()
    print("=== REX COMPARISON ===")
    if r:
        for k in r.keys():
            if r[k] != 0 and r[k] != "" and r[k] is not None:
                print(f"  {k}: {r[k]}")
                
    # Query liquidaciones
    cursor.execute("SELECT * FROM liquidaciones WHERE rut = '17773864-6' AND periodo = '2026-05'")
    l = cursor.fetchone()
    print("\n=== CALCULATED LIQUIDACION ===")
    if l:
        for k in l.keys():
            if l[k] != 0 and l[k] != "" and l[k] is not None:
                print(f"  {k}: {l[k]}")
                
    conn.close()

if __name__ == "__main__":
    inspect()

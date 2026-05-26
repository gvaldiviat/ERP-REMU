import sqlite3

DB_PATH = "remuneraciones.db"
periodo = "2026-05"

def check():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT DISTINCT snapshot_timestamp 
        FROM liquidaciones_snapshots 
        WHERE periodo = ?
        ORDER BY snapshot_timestamp DESC
    """, (periodo,))
    timestamps = [r[0] for r in cursor.fetchall() if r[0]]
    print("Timestamps list:", timestamps)
    
    latest_ts = timestamps[0]
    
    if len(timestamps) > 1:
        for ts in timestamps:
            cursor.execute("SELECT COUNT(*) FROM liquidaciones_snapshots WHERE periodo = ? AND snapshot_timestamp = ?", (periodo, ts))
            snap_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM liquidaciones WHERE periodo = ?", (periodo,))
            curr_count = cursor.fetchone()[0]
            
            print(f"Checking ts: {ts} -> snap_count={snap_count}, curr_count={curr_count}")
            
            if snap_count != curr_count:
                print(f"  -> Count differs! Setting latest_ts = {ts}")
                latest_ts = ts
                break
                
            cursor.execute("SELECT SUM(costo_empresa) FROM liquidaciones_snapshots WHERE periodo = ? AND snapshot_timestamp = ?", (periodo, ts))
            snap_cost = cursor.fetchone()[0] or 0
            
            cursor.execute("SELECT SUM(costo_empresa) FROM liquidaciones WHERE periodo = ?", (periodo,))
            curr_cost = cursor.fetchone()[0] or 0
            
            print(f"  -> snap_cost={snap_cost}, curr_cost={curr_cost}, diff={abs(snap_cost - curr_cost)}")
            
            if abs(snap_cost - curr_cost) > 10:
                print(f"  -> Cost differs! Setting latest_ts = {ts}")
                latest_ts = ts
                break
                
    print("Selected latest_ts:", latest_ts)
    conn.close()

if __name__ == "__main__":
    check()

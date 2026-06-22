#!/usr/bin/env python3
"""
ERP REMU - Excel Exporter
Generates a beautifully formatted analytical spreadsheet with pivot-friendly sheets.
"""

from exporters import generate_excel
import os
import sys

def main():
    periodo = "2026-05"
    if len(sys.argv) > 1:
        periodo = sys.argv[1]

    print("=======================================================")
    print(f"[*] ERP REMU - Generando Reporte Excel Analítico para el Periodo {periodo}...")
    print("=======================================================")
    
    try:
        file_bytes = generate_excel(periodo)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(base_dir, f"Reporte_Analitico_Remuneraciones_{periodo}.xlsx")
        
        with open(out_path, "wb") as f:
            f.write(file_bytes)
            
        print(f"\n[OK] EXITO: Archivo Excel generado exitosamente!")
        print(f"[*] Ubicación: {out_path}")
        print(f"[*] Tamaño: {round(os.path.getsize(out_path) / 1024, 2)} KB")
        print("=======================================================\n")
    except Exception as e:
        print(f"\n[ERROR] ERROR: No se pudo generar el archivo Excel.")
        print(f"    Detalle del error: {str(e)}")
        print("=======================================================\n")

if __name__ == "__main__":
    main()

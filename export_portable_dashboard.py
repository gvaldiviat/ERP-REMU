#!/usr/bin/env python3
"""
ERP REMU - Offline Dashboard Exporter
Generates a fully self-contained, interactive HTML file that runs 100% offline.
This file can be uploaded directly to OneDrive, SharePoint, or shared via email.
"""

from exporters import generate_portable_dashboard
import os

def main():
    print("=======================================================")
    print("[*] ERP REMU - Generando Panel de Control Portable Offline...")
    print("=======================================================")
    
    try:
        compiled_path = generate_portable_dashboard()
        print(f"\n[OK] EXITO: Panel de control portable generado exitosamente!")
        print(f"[*] Ubicación: {compiled_path}")
        print(f"[*] Tamaño: {round(os.path.getsize(compiled_path) / 1024 / 1024, 2)} MB")
        print("\n[i] Instrucción de Uso:")
        print("    1. Copia o mueve el archivo 'dashboard_portable.html' a tu OneDrive o SharePoint.")
        print("    2. Comparte el enlace de OneDrive/SharePoint con tu equipo o clientes.")
        print("    3. Podrán abrir y usar la aplicación completa de forma 100% interactiva en cualquier navegador.")
        print("=======================================================\n")
    except Exception as e:
        print(f"\n[ERROR] ERROR: No se pudo compilar el dashboard portable.")
        print(f"    Detalle del error: {str(e)}")
        print("=======================================================\n")

if __name__ == "__main__":
    main()

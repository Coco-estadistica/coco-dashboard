#!/usr/bin/env python3
"""
Sincroniza index.html y Dashboard_COCO.html para evitar desincronización.
Ejecuta esto después de editar cualquier HTML para mantener ambos en sync.
"""

import shutil
from datetime import datetime

def sincronizar():
    archivo_origen = "Dashboard_COCO.html"
    archivo_destino = "index.html"

    print(f"Sincronizando {archivo_origen} → {archivo_destino}...")
    shutil.copy(archivo_origen, archivo_destino)
    print(f"✓ Sincronización completada: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Ambos archivos tienen los mismos cambios.")

if __name__ == "__main__":
    sincronizar()

#!/usr/bin/env python3
"""
Script para cargar datos de análisis MRR a la BD_MAESTRA_COCO.xlsx
Carga dos hojas: Analisis_MRR_Resumen y Analisis_MRR_Detalle
"""

import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
import shutil

def cargar_datos_mrr():
    """Carga datos de MRR desde archivo de análisis a la base de datos maestra"""

    # Rutas
    archivo_mrr = 'Analisis_MRR_Junio_Julio_2.xlsx'
    archivo_bd = 'migracion/BD_MAESTRA_COCO.xlsx'
    archivo_respaldo = f'migracion/BD_MAESTRA_COCO_respaldo_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

    print(f"📋 Cargando análisis MRR desde: {archivo_mrr}")
    print(f"📊 Base de datos destino: {archivo_bd}\n")

    # Crear respaldo
    print("✓ Creando respaldo de la base de datos...")
    shutil.copy(archivo_bd, archivo_respaldo)
    print(f"  Respaldo guardado en: {archivo_respaldo}\n")

    # Leer archivo MRR
    print("✓ Leyendo archivo MRR...")
    wb_mrr = openpyxl.load_workbook(archivo_mrr)

    # Leer base de datos
    print("✓ Leyendo base de datos maestra...")
    wb_bd = openpyxl.load_workbook(archivo_bd)

    # Procesar cada hoja del archivo MRR
    for sheet_name in wb_mrr.sheetnames:
        print(f"\n✓ Procesando hoja: {sheet_name}")
        ws_mrr = wb_mrr[sheet_name]

        # Crear nombre para la hoja destino
        nombre_destino = f"Analisis_MRR_{sheet_name.replace(' ', '_')}"

        # Eliminar hoja destino si existe
        if nombre_destino in wb_bd.sheetnames:
            print(f"  - Eliminando hoja existente: {nombre_destino}")
            del wb_bd[nombre_destino]

        # Crear nueva hoja
        ws_destino = wb_bd.create_sheet(nombre_destino)
        print(f"  - Creada nueva hoja: {nombre_destino}")

        # Copiar datos
        filas_copiadas = 0
        for row_idx, row in enumerate(ws_mrr.iter_rows(), 1):
            for col_idx, cell in enumerate(row, 1):
                nueva_celda = ws_destino.cell(row=row_idx, column=col_idx)

                # Copiar valor
                if cell.value is not None:
                    nueva_celda.value = cell.value
                    filas_copiadas += 1

                # Copiar formato
                if cell.has_style:
                    nueva_celda.font = openpyxl.styles.Font(
                        name=cell.font.name,
                        size=cell.font.size,
                        bold=cell.font.bold,
                        italic=cell.font.italic,
                        color=cell.font.color
                    )
                    nueva_celda.fill = openpyxl.styles.PatternFill(
                        fill_type=cell.fill.fill_type,
                        start_color=cell.fill.start_color,
                        end_color=cell.fill.end_color
                    )
                    nueva_celda.alignment = openpyxl.styles.Alignment(
                        horizontal=cell.alignment.horizontal,
                        vertical=cell.alignment.vertical,
                        wrap_text=cell.alignment.wrap_text
                    )

        # Ajustar ancho de columnas
        for column in ws_destino.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws_destino.column_dimensions[column_letter].width = adjusted_width

        print(f"  - Copiadas {filas_copiadas} celdas con datos")

    # Guardar base de datos actualizada
    print(f"\n✓ Guardando base de datos actualizada...")
    wb_bd.save(archivo_bd)
    print(f"✓ Guardado: {archivo_bd}\n")

    # Resumen
    print("=" * 60)
    print("📈 CARGA COMPLETADA")
    print("=" * 60)
    print(f"Hojas cargadas:")
    for sheet_name in wb_mrr.sheetnames:
        nombre_destino = f"Analisis_MRR_{sheet_name.replace(' ', '_')}"
        print(f"  ✓ {nombre_destino}")

    print(f"\n📌 Notas:")
    print(f"  - Respaldo disponible en: {archivo_respaldo}")
    print(f"  - La base de datos está lista para usar")
    print(f"  - Abre 'Abrir_Dashboard_COCO.bat' para ver los cambios")
    print()

if __name__ == "__main__":
    cargar_datos_mrr()

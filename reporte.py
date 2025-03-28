import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

def generar_reporte_excel(json_file, excel_file):
    # Cargar datos desde el archivo JSON
    with open(json_file, 'r', encoding='utf-8') as file:
        datos = json.load(file)

    # Crear un libro de Excel y una hoja
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registro"

    # Agregar título al reporte
    ws.merge_cells("A1:H1")
    titulo = ws.cell(row=1, column=1, value="Reportes")  # Título del reporte
    titulo.font = Font(bold=True, size=14)
    titulo.alignment = Alignment(horizontal='center')  # Centrar el título

    # Definir bordes
    borde = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )

    # Definir encabezados con columna de correlativo
    encabezados = ["N°", "Número de DNI", "Apellidos y Nombres", "Correo", "Celular", "Sede", "Programa", "Fecha de Registro"]
    ws.append(encabezados)

    # Aplicar estilos a los encabezados
    for col_num, col_name in enumerate(encabezados, start=1):
        cell = ws.cell(row=2, column=col_num, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.border = borde  # Aplicar bordes

    # Llenar los datos en la hoja con el correlativo
    for index, persona in enumerate(datos, start=1):
        row = [
            index,  # Correlativo
            persona.get("dni", ""),
            persona.get("nombre", ""),
            persona.get("correo", ""),
            persona.get("celular", ""),
            persona.get("sede", ""),
            persona.get("programa", ""),
            persona.get("fecha_registro", "")
        ]
        ws.append(row)

        # Aplicar bordes a cada celda
        for col_num in range(1, len(row) + 1):
            cell = ws.cell(row=index + 2, column=col_num)
            cell.border = borde  # Bordes en todas las celdas

        # Hacer que los números de la columna "N°" sean negros suaves
        ws.cell(row=index + 2, column=1).font = Font(size=10, bold=True, color="202020")  # Negro suave
        ws.cell(row=index + 2, column=1).alignment = Alignment(horizontal='center')  # Centrar correlativo

    # Ajustar el ancho de columnas automáticamente
    for col in ws.iter_cols(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        max_length = 0
        col_letter = col[0].column_letter  # Obtener la letra de la columna
        for cell in col:
            if cell.value and isinstance(cell, openpyxl.cell.cell.Cell):  # Evitar MergedCell
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    # Guardar el archivo Excel
    wb.save(excel_file)
    print(f"Reporte generado con éxito: {excel_file}")

# Ejemplo de uso
generar_reporte_excel("registro.json", "reporte_registro.xlsx")

#python reporte.py
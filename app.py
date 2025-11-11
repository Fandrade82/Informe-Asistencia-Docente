import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from flask import Flask, request, send_file, render_template
import io
from datetime import datetime

app = Flask(__name__)
processed_data = None

# --- Función para generar Excel con formato ---
def generar_excel(df):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Eliminar hoja por defecto

    # Estilos
    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    border_style = Border(left=Side(style="thin"), right=Side(style="thin"),
                           top=Side(style="thin"), bottom=Side(style="thin"))
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Agrupar por docente
    for docente, grupo in df.groupby("Apellido y Nombre"):
        ws = wb.create_sheet(title=docente[:30])  # Nombre hoja limitado a 30 caracteres

        # Datos del docente
        codigo = grupo["Codigo"].iloc[0]
        departamento = grupo["Departamento"].iloc[0]
        mes = grupo["Fecha"].dt.strftime("%B %Y").iloc[0].upper()

        # Título
        titulo = f"REPORTE DE ASISTENCIA - {docente} | Código: {codigo} | Departamento: {departamento} | Mes: {mes}"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
        cell_title = ws.cell(row=1, column=1, value=titulo)
        cell_title.font = Font(bold=True, size=14)
        cell_title.alignment = center_align

        # Encabezados
        headers = ["Fecha", "Semana", "Hora1", "Hora2", "Hora3", "Hora4", "Tiempo total trabajado", "Observaciones"]
        ws.append(headers)
        for col in range(1, len(headers)+1):
            cell = ws.cell(row=2, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border_style

        # Filas de datos
        total_trabajado = 0
        for _, row in grupo.iterrows():
            fecha = row["Fecha"].strftime("%d/%m/%Y")
            semana = row["Semana"]
            hora1 = row.get("Hora1", "")
            hora2 = row.get("Hora2", "")
            hora3 = row.get("Hora3", "")
            hora4 = row.get("Hora4", "")
            tiempo = row.get("Tiempo total de trabajo", 0)
            observacion = ""

            # Validación de marcación
            if pd.isna(hora3) or pd.isna(hora4):
                observacion = "No marca"

            ws.append([fecha, semana, hora1, hora2, hora3, hora4, tiempo, observacion])
            total_trabajado += tiempo

            # Formato fila
            for col in range(1, 9):
                cell = ws.cell(row=ws.max_row, column=col)
                cell.border = border_style
                cell.alignment = center_align

            # Resaltar Hora3 y Hora4 si falta marcación
            if observacion == "No marca":
                ws.cell(row=ws.max_row, column=5).fill = yellow_fill
                ws.cell(row=ws.max_row, column=6).fill = yellow_fill

        # Fila total
        ws.append(["", "", "", "", "", "Total trabajado", total_trabajado, ""])
        for col in range(1, 9):
            cell = ws.cell(row=ws.max_row, column=col)
            cell.border = border_style
            cell.alignment = center_align
            if col == 6:
                cell.font = Font(bold=True)

        # Ajustar ancho de columnas
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # Guardar en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html', ready=False)

@app.route('/upload', methods=['POST'])
def upload_file():
    global processed_data
    file = request.files['file']
    df = pd.read_excel(file)
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    processed_data = df
    return render_template('index.html', ready=True)

@app.route('/download_excel')
def download_excel():
    output = generar_excel(processed_data)
    return send_file(output, as_attachment=True, download_name="Informe_Asistencia.xlsx")

if __name__ == '__main__':
    app.run(debug=True)

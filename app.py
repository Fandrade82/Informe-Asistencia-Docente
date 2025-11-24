from flask import Flask, request, send_file, render_template
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')  # Usa el HTML con Bootstrap en /templates

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        # Leer archivo subido
        file = request.files['file']
        df = pd.read_excel(file, engine="openpyxl")

        # Validar columnas requeridas
        required_cols = [
            "Fecha", "Semana", "Hora1", "Hora2", "Hora3", "Hora4",
            "Tiempo total de trabajo", "Departamento", "Apellido y Nombre"
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return f"Error: Faltan columnas en el archivo: {', '.join(missing)}", 400

        # Eliminar sábados y domingos
        df = df[~df["Semana"].isin(["sábado", "domingo"])]

        # Crear libro de salida
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Asistencia"

        # Estilos
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_align = Alignment(horizontal="center", vertical="center")

        # Agrupar por docente
        grouped = df.groupby("Apellido y Nombre")

        row = 1
        for docente, data in grouped:
            # Título con nombre del docente
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            cell = ws.cell(row=row, column=1, value=docente)
            cell.font = Font(bold=True, size=14)
            cell.alignment = center_align
            row += 1

            # Encabezado
            headers = ["Fecha", "Departamento", "Hora1", "Hora2", "Hora3", "Hora4", "Tiempo total", "Observación"]
            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col_idx, value=h)

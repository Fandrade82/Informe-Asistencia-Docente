from flask import Flask, request, send_file, render_template
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            df = pd.read_excel(file, engine='openpyxl')
            df.columns = df.columns.str.strip()

            docentes = df['Apellido y Nombre'].unique()
            wb = Workbook()
            ws = wb.active
            ws.title = "Informe"

            header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            center_align = Alignment(horizontal="center")

            row_num = 1
            for docente in docentes:
                ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=10)
                cell = ws.cell(row=row_num, column=1, value=docente)
                cell.font = Font(bold=True)
                cell.alignment = center_align
                row_num += 1

                subset = df[df['Apellido y Nombre'] == docente]
                columns = list(subset.columns) + ['Observación']
                for col_num, col_name in enumerate(columns, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=col_name)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align
                row_num += 1

                for _, row in subset.iterrows():
                    jornada = str(row['Departamento']).strip().upper()
                    dia_semana = str(row['Semana']).strip().lower()
                    verificar = False
                    if jornada.startswith('ADMIN'):
                        verificar = True
                    elif jornada.startswith('DOC') and dia_semana == 'jueves':
                        verificar = True

                    observacion = ''
                    if verificar and (pd.isna(row.get('Hora3')) or pd.isna(row.get('Hora4'))):
                        observacion = 'No marca'

                    for col_num, col_name in enumerate(subset.columns, 1):
                        valor = row.get(col_name)
                        cell = ws.cell(row=row_num, column=col_num, value=valor)
                        if observacion == 'No marca':
                            cell.fill = yellow_fill
                        cell.alignment = center_align

                    cell = ws.cell(row=row_num, column=len(subset.columns)+1, value=observacion)
                    if observacion == 'No marca':
                        cell.fill = yellow_fill
                    cell.alignment = center_align
                    row_num += 1

                row_num += 1  # Espacio entre docentes

            output_path = "informe_asistencia.xlsx"
            wb.save(output_path)
            return send_file(output_path, as_attachment=True)

    return render_template("index.html")

from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            # Leer el Excel original
            df = pd.read_excel(file)

            # Procesamiento según reglas:
            # 1. Consolidar datos por docente (ejemplo: agrupar por 'Docente')
            if 'Docente' in df.columns:
                grouped = df.groupby('Docente').agg(lambda x: ' | '.join(map(str, x)))
            else:
                grouped = df  # Si no hay columna Docente, usar tal cual

            # Crear nuevo Excel con formato
            wb = Workbook()
            ws = wb.active
            ws.title = "Informe Asistencia"

            # Encabezado azul con letras blancas
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            alignment_center = Alignment(horizontal="center", vertical="center")

            # Escribir encabezados
            headers = list(grouped.columns)
            ws.append(["Docente"] + headers)
            for col in range(1, len(headers) + 2):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = alignment_center

            # Escribir datos y aplicar reglas
            for idx, (docente, row) in enumerate(grouped.iterrows(), start=2):
                ws.cell(row=idx, column=1, value=docente)
                for col_idx, value in enumerate(row, start=2):
                    cell = ws.cell(row=idx, column=col_idx, value=value)
                    # Si falta marcación, resaltar en amarillo y poner "No marca"
                    if pd.isna(value) or str(value).strip() == "":
                        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        cell.value = "No marca"

            # Ajustar ancho de columnas
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_length + 2

            # Guardar en memoria
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            return send_file(output,
                             as_attachment=True,
                             download_name="informe_asistencia.xlsx",
                             mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from flask import Flask, request, send_file, render_template
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)

# Variable global para almacenar datos procesados
processed_data = None

# --- Función para generar Excel ---
def generar_excel(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Informe Asistencia"

    # Estilos encabezado
    header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center")

    # Encabezados
    headers = ["Docente", "Fecha", "Hora3", "Hora4", "Observación"]
    ws.append(headers)
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Procesar filas
    for _, row in df.iterrows():
        docente = row["Docente"]
        fecha = row["Fecha"]
        hora3 = row.get("Hora3", "")
        hora4 = row.get("Hora4", "")
        jornada = row.get("Jornada", "")
        observacion = ""

        # Validaciones
        if jornada in ["ADMIN MATUTINA", "ADMIN VESPERTINA"]:
            if pd.isna(hora3) or pd.isna(hora4):
                observacion = "No marca"
        elif jornada in ["DOC. MATUTINA", "DOC. VESPERTINA", "DOC. NOCTURNA"]:
            if fecha.weekday() == 3:  # Jueves
                if pd.isna(hora3) or pd.isna(hora4):
                    observacion = "No marca"

        ws.append([docente, fecha.strftime("%d/%m/%Y"), hora3, hora4, observacion])

        # Resaltar si falta marcación
        if observacion == "No marca":
            for col in range(1, 6):
                ws.cell(row=ws.max_row, column=col).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Ajustar ancho de columnas
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # Guardar en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Función para generar PDF ---
def generar_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(200, 770, "Informe de Asistencia")
    c.setFont("Helvetica", 10)
    y = 740
    for _, row in df.iterrows():
        texto = f"{row['Docente']} | {row['Fecha'].strftime('%d/%m/%Y')} | Hora3: {row.get('Hora3','')} | Hora4: {row.get('Hora4','')}"
        c.drawString(50, y, texto)
        y -= 15
        if y < 50:
            c.showPage()
            y = 750
    c.save()
    buffer.seek(0)
    return buffer

# --- Rutas ---
@app.route('/')
def index():
    return render_template('index.html', ready=False)

@app.route('/upload', methods=['POST'])
def upload_file():
    global processed_data
    file = request.files['file']
    df = pd.read_excel(file)
    processed_data = df
    return render_template('index.html', ready=True)

@app.route('/download_excel')
def download_excel():
    output = generar_excel(processed_data)
    return send_file(output, as_attachment=True, download_name="Informe_Asistencia.xlsx")

@app.route('/download_pdf')
def download_pdf():
    output = generar_pdf(processed_data)
    return send_file(output, as_attachment=True, download_name="Informe_Asistencia.pdf")

if __name__ == '__main__':
    app.run(debug=True)

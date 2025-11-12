
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from flask import Flask, request, send_file, render_template
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)

def validar_columnas(df):
    required_cols = ["Docente", "Fecha", "Hora3", "Hora4", "Jornada"]
    return all(col in df.columns for col in required_cols)

def generar_excel(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Informe Asistencia"

    header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center")

    headers = ["Docente", "Fecha", "Hora3", "Hora4", "Observación"]
    ws.append(headers)
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    for _, row in df.iterrows():
        docente = row.get("Docente", "")
        fecha = row.get("Fecha", "")
        hora3 = row.get("Hora3", "") or ""
        hora4 = row.get("Hora4", "") or ""
        jornada = row.get("Jornada", "")
        observacion = ""

        if isinstance(fecha, pd.Timestamp):
            if jornada in ["ADMIN MATUTINA", "ADMIN VESPERTINA"]:
                if pd.isna(hora3) or pd.isna(hora4):
                    observacion = "No marca"
            elif jornada in ["DOC. MATUTINA", "DOC. VESPERTINA", "DOC. NOCTURNA"]:
                if fecha.weekday() == 3:
                    if pd.isna(hora3) or pd.isna(hora4):
                        observacion = "No marca"

        fecha_str = fecha.strftime("%d/%m/%Y") if isinstance(fecha, pd.Timestamp) else str(fecha)
        ws.append([docente, fecha_str, hora3, hora4, observacion])

        if observacion == "No marca":
            for col in range(1, 6):
                ws.cell(row=ws.max_row, column=col).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generar_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(200, 770, "Informe de Asistencia")
    c.setFont("Helvetica", 10)
    y = 740
    for _, row in df.iterrows():
        fecha_str = row["Fecha"].strftime("%d/%m/%Y") if isinstance(row["Fecha"], pd.Timestamp) else str(row["Fecha"])
        texto = f"{row.get('Docente','')} | {fecha_str} | Hora3: {row.get('Hora3','')} | Hora4: {row.get('Hora4','')}"
        c.drawString(50, y, texto)
        y -= 15
        if y < 50:
            c.showPage()
            y = 750
    c.save()
    buffer.seek(0)
    return buffer

@app.route('/')
def index():
    return render_template('index.html', ready=False)

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return "Error: Formato no permitido"
    df = pd.read_excel(file)
    if not validar_columnas(df):
        return "Error: El archivo no tiene las columnas requeridas"
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    request.environ['processed_data'] = df
    return render_template('index.html', ready=True)

@app.route('/download_excel')
def download_excel():
    df = request.environ.get('processed_data')
    if df is None:
        return "Error: No hay datos procesados"
    output = generar_excel(df)
    return send_file(output, as_attachment=True, download_name="Informe_Asistencia.xlsx")

@app.route('/download_pdf')
def download_pdf():
    df = request.environ.get('processed_data')
    if df is None:
        return "Error: No hay datos procesados"
    output = generar_pdf(df)
    return send_file(output, as_attachment=True, download_name="Informe_Asistencia.pdf")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

# ============================================
# app.py - Sistema de Informe de Asistencia
# ============================================

from flask import Flask, request, send_file, render_template, flash, redirect, url_for
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import datetime
import os
import tempfile

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui_cambiar_en_produccion'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Columnas requeridas en el Excel
REQUIRED_COLUMNS = ['Apellido y Nombre', 'Departamento', 'Semana']

def validate_excel(df):
    """Valida que el DataFrame tenga las columnas necesarias"""
    df.columns = df.columns.str.strip()
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        return False, f"Faltan columnas requeridas: {', '.join(missing)}"
    return True, "OK"

def generate_report(df):
    """Genera el reporte Excel con formato"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Informe"
    
    # Estilos
    header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    docentes = df['Apellido y Nombre'].unique()
    row_num = 1
    
    for docente in docentes:
        # Encabezado del docente
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=10)
        cell = ws.cell(row=row_num, column=1, value=docente)
        cell.font = Font(bold=True, size=12)
        cell.alignment = center_align
        cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        row_num += 1
        
        # Filtrar datos del docente
        subset = df[df['Apellido y Nombre'] == docente].copy()
        columns = list(subset.columns) + ['Observación']
        
        # Encabezados de columnas
        for col_num, col_name in enumerate(columns, 1):
            cell = ws.cell(row=row_num, column=col_num, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
        row_num += 1
        
        # Datos del docente
        for _, row in subset.iterrows():
            jornada = str(row.get('Departamento', '')).strip().upper()
            dia_semana = str(row.get('Semana', '')).strip().lower()
            
            # Lógica de verificación
            verificar = False
            if jornada.startswith('ADMIN'):
                verificar = True
            elif jornada.startswith('DOC') and dia_semana == 'jueves':
                verificar = True
            
            # Verificar marcaciones
            observacion = ''
            hora3 = row.get('Hora3')
            hora4 = row.get('Hora4')
            
            if verificar:
                hora3_vacia = pd.isna(hora3) or (isinstance(hora3, str) and not hora3.strip())
                hora4_vacia = pd.isna(hora4) or (isinstance(hora4, str) and not hora4.strip())
                
                if hora3_vacia or hora4_vacia:
                    observacion = 'No marca'
            
            # Escribir datos de la fila
            for col_num, col_name in enumerate(subset.columns, 1):
                valor = row.get(col_name)
                if pd.isna(valor):
                    valor = ''
                
                cell = ws.cell(row=row_num, column=col_num, value=valor)
                if observacion == 'No marca':
                    cell.fill = yellow_fill
                cell.alignment = center_align
            
            # Columna de observación
            cell = ws.cell(row=row_num, column=len(subset.columns)+1, value=observacion)
            if observacion == 'No marca':
                cell.fill = yellow_fill
                cell.font = Font(bold=True)
            cell.alignment = center_align
            row_num += 1
        
        row_num += 1
    
    # Ajustar ancho de columnas
    for column in ws.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    
    return wb

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Verificar que se subió un archivo
        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        
        # Verificar que el archivo tiene nombre
        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
        
        # Verificar extensión
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            flash('Solo se permiten archivos Excel (.xlsx, .xls)', 'error')
            return redirect(request.url)
        
        try:
            # Leer el archivo Excel
            df = pd.read_excel(file, engine='openpyxl')
            
            # Verificar que el DataFrame no esté vacío
            if df.empty:
                flash('El archivo Excel está vacío', 'error')
                return redirect(request.url)
            
            # Validar columnas
            valid, message = validate_excel(df)
            if not valid:
                flash(message, 'error')
                return redirect(request.url)
            
            # Generar reporte
            wb = generate_report(df)
            
            # Guardar en archivo temporal
            fecha = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            temp_dir = tempfile.gettempdir()
            output_path = os.path.join(temp_dir, f"Informe_Asistencia_{fecha}.xlsx")
            wb.save(output_path)
            
            # Enviar archivo y programar eliminación
            response = send_file(
                output_path, 
                as_attachment=True,
                download_name=f"Informe_Asistencia_{fecha}.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # Eliminar archivo después de enviarlo
            @response.call_on_close
            def cleanup():
                try:
                    if os.path.exists(output_path):
                        os.remove(output_path)
                        print(f"✓ Archivo temporal eliminado: {output_path}")
                except Exception as e:
                    print(f"⚠ Error al eliminar archivo temporal: {e}")
            
            return response
            
        except pd.errors.EmptyDataError:
            flash('El archivo Excel está vacío o corrupto', 'error')
            return redirect(request.url)
        except Exception as e:
            flash(f'Error al procesar el archivo: {str(e)}', 'error')
            print(f"Error detallado: {e}")
            return redirect(request.url)
    
    return render_template('index.html')

@app.errorhandler(413)
def too_large(e):
    flash('El archivo es demasiado grande. Máximo permitido: 16MB', 'error')
    return redirect(url_for('upload_file'))

@app.errorhandler(500)
def internal_error(e):
    flash('Error interno del servidor. Por favor intenta nuevamente.', 'error')
    return redirect(url_for('upload_file'))

if __name__ == '__main__':
    print("=" * 60)
    print("🚀 Servidor iniciado correctamente")
    print("📍 URL: http://127.0.0.1:5000")
    print("📁 Asegúrate de tener la carpeta 'templates' con index.html")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)

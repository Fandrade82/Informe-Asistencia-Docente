# Informe de Asistencia Docente

Aplicación web en Flask para generar informes de asistencia docente a partir de archivos Excel.

## ✅ Funcionalidades
- Subir archivo Excel con datos de asistencia.
- Procesar datos según reglas:
  - Verificación de Hora3 y Hora4 según jornada y día.
  - Observación "No marca" si falta marcación.
  - Resaltado en amarillo para filas incompletas.
- Descargar informe final en formato Excel.

## 📦 Requisitos
- Python 3.x
- Flask
- Pandas
- OpenPyXL
- Gunicorn (para despliegue en Render)

## 🚀 Cómo ejecutar localmente
```bash
pip install -r requirements.txt
python3 app.py

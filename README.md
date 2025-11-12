
# Proyecto Informe de Asistencia

## Cómo ejecutar localmente
1. Crear entorno virtual:
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/Mac
   venv\Scripts\activate   # Windows
   ```
2. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Ejecutar la aplicación:
   ```bash
   python app.py
   ```
4. Abrir en el navegador:
   ```
   http://127.0.0.1:5000
   ```

## Despliegue en Render
- Render detectará `render.yaml` y ejecutará:
  ```
  gunicorn app:app --bind 0.0.0.0:$PORT
  ```

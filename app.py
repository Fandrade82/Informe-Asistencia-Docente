
from flask import Flask, render_template, request
import pandas as pd
import openpyxl

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            df = pd.read_excel(file)
            return f"<h3>Archivo procesado correctamente. Filas: {len(df)}</h3>"
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

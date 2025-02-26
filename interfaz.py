from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Ruta principal para cargar el archivo
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Verificar si se cargó un archivo
        if "inventario_file" not in request.files:
            return "No se cargó ningún archivo", 400
        
        file = request.files["inventario_file"]
        if file.filename == "":
            return "Nombre de archivo inválido", 400
        
        # Guardar el archivo cargado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        # Procesar el archivo (usar tu código actual)
        try:
            # Cargar el archivo de inventario
            inventario = pd.read_excel(file_path)
            inventario["UPC"] = inventario["UPC"].astype(str)
            inventario["UPC"] = inventario["UPC"].str.replace(".0", "")
            
            # Aquí puedes agregar el resto de tu código para procesar el inventario
            # y generar los archivos de salida.
            
            # Ejemplo: Guardar un archivo de salida
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], "output.xlsx")
            inventario.to_excel(output_path, index=False)
            
            # Devolver el archivo generado para descargar
            return send_file(output_path, as_attachment=True)
        
        except Exception as e:
            return f"Error al procesar el archivo: {str(e)}", 500
    
    # Mostrar el formulario de carga (GET)
    return render_template("index.html")

# Iniciar la aplicación
if __name__ == "__main__":
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True)
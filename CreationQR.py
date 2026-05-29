import os
import io
import zipfile
import pandas as pd
import qrcode
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

app = Flask(__name__)
app.secret_key = "qr_secret_key"  # Necesario para mostrar mensajes de error/éxito

def generate_single_qr(data):
    """Genera un objeto de imagen QR en memoria."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")

def clean_filename(name):
    """Limpia el nombre del archivo para evitar caracteres inválidos."""
    return "".join(c for c in str(name) if c.isalnum() or c in (' ', '_', '-')).rstrip()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # ==========================================
        # MÓDULO INDIVIDUAL
        # ==========================================
        if 'btn_individual' in request.form:
            link = request.form.get('link')
            nombre = request.form.get('nombre') or "codigo_qr"
            
            if not link:
                flash("Por favor, ingresa un link.", "danger")
                return redirect(url_for('index'))

            img = generate_single_qr(link)
            img_io = io.BytesIO()
            img.save(img_io, 'PNG')
            img_io.seek(0)
            
            return send_file(img_io, mimetype='image/png', as_attachment=True, download_name=f"{clean_filename(nombre)}.png")

        # ==========================================
        # MÓDULO MASIVO (EXCEL)
        # ==========================================
        elif 'btn_masivo_submit' in request.form:
            if 'file_excel' not in request.files:
                flash("No se encontró el campo de archivo.", "danger")
                return redirect(url_for('index'))
                
            file = request.files['file_excel']
            if file.filename == '':
                flash("No seleccionaste ningún archivo Excel.", "danger")
                return redirect(url_for('index'))

            try:
                # Se lee el archivo asegurando compatibilidad con .xlsx
                df = pd.read_excel(file, engine='openpyxl')
                
                # --- NUEVA VALIDACIÓN DE CABECERAS EXCLUSIVAS (QA BLINDAJE) ---
                # Definimos los nombres exactos que el usuario lee en el banner informativo del HTML
                col_id = 'ID (Nombre del archivo)'
                col_link = 'Link (URL del QR)'
                
                # Si las columnas requeridas no existen en el Excel, frenamos el proceso de inmediato
                if col_id not in df.columns or col_link not in df.columns:
                    flash("Estructura incorrecta: El archivo Excel debe tener exactamente las columnas 'ID (Nombre del archivo)' y 'Link (URL del QR)'. Verifique mayúsculas y acentos.", "danger")
                    return redirect(url_for('index'))
                
                # --- VALIDACIÓN DE EXCEL VACÍO ---
                if df.empty:
                    flash("El archivo Excel seleccionado está vacío.", "warning")
                    return redirect(url_for('index'))

                zip_io = io.BytesIO()
                qr_generados = 0
                
                with zipfile.ZipFile(zip_io, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for i, row in df.iterrows():
                        # Extraemos la información usando las llaves de columna exactas y validadas
                        link = str(row.get(col_link, '')).strip()
                        nombre_id = str(row.get(col_id, '')).strip()

                        # Ignorar explícitamente celdas vacías o nulas de Pandas (NaN)
                        if not link or link.lower() == 'nan' or link == '':
                            continue
                        
                        # Si el ID está vacío, le asignamos un correlativo por defecto para que no falle
                        if not nombre_id or nombre_id.lower() == 'nan':
                            nombre_id = f'QR_{i+1}'

                        img = generate_single_qr(link)
                        img_io = io.BytesIO()
                        img.save(img_io, 'PNG')
                        img_io.seek(0)
                        
                        filename = f"{clean_filename(nombre_id)}.png"
                        zip_file.writestr(filename, img_io.getvalue())
                        qr_generados += 1
                
                # Si el bucle terminó pero no se procesó ningún link válido
                if qr_generados == 0:
                    flash("No se encontraron enlaces válidos para procesar dentro del archivo.", "warning")
                    return redirect(url_for('index'))
                
                zip_io.seek(0)
                return send_file(zip_io, mimetype='application/zip', as_attachment=True, download_name="QRs_Masivos.zip")

            except Exception as e:
                flash(f"Error crítico al procesar el Excel: {str(e)}", "danger")
                return redirect(url_for('index'))

    return render_template('index.html')


@app.route('/descargar-plantilla')
def descargar_plantilla():
    """Genera y descarga dinámicamente un Excel plantilla con las cabeceras requeridas."""
    try:
        # Definimos las columnas exactas que la app leerá y validará en el módulo masivo
        columnas = ['N°', 'ID (Nombre del archivo)', 'Link (URL del QR)']
        
        # Creamos un DataFrame vacío estructurado solo con las cabeceras
        df_plantilla = pd.DataFrame(columns=columnas)
        
        # Guardamos el Excel en un buffer binario en la memoria RAM
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_plantilla.to_excel(writer, index=False, sheet_name='Plantilla_QR')
        
        output.seek(0)
        
        # Entregamos el archivo listo para descarga directa
        return send_file(
            output, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            as_attachment=True, 
            download_name="Plantilla_Masivo_QR.xlsx"
        )
    except Exception as e:
        flash(f"No se pudo generar el archivo de plantilla: {str(e)}", "danger")
        return redirect(url_for('index'))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
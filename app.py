from flask import Flask, render_template, request, send_file
from excel_utils import leer_items_y_descripciones_excel
from pdf_utils import generar_pdf
from match_utils import aplicar_match_a_filas
from pathlib import Path
import time

app = Flask(__name__)

UPLOADS = Path("uploads")
UPLOADS.mkdir(exist_ok=True)

# Archivo de coincidencias (match.xlsx)
# Ubicarlo en la misma carpeta que app.py (pdf_generator/match.xlsx)
MATCH_XLSX = Path(__file__).resolve().parent / "match.xlsx"

# Logo por defecto (se usa si el usuario no sube uno)
DEFAULT_LOGO = Path(__file__).resolve().parent / "logo_default.png"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        # =============================
        # FECHA INGRESADA EN EL FORM
        # =============================
        fecha = request.form.get("fecha", "").strip()

        # =============================
        # TEMPLATE (selector)
        # =============================
        template_id = (request.form.get("template", "1") or "1").strip()
        if template_id not in ("1", "2"):
            template_id = "1"

        # Validación: si se elige Template 2, debe existir el PDF en el proyecto
        if template_id == "2":
            t2 = Path(__file__).resolve().parent / "template_desglose2.pdf"
            if not t2.exists():
                return "No se encontró template_desglose2.pdf en el servidor. Agregalo en la carpeta pdf_generator/ o elige Template 1.", 400

        # =============================
        # EXCEL SUBIDO
        # =============================
        archivo = request.files.get("excel")
        if not archivo:
            return "No se subió ningún archivo Excel", 400

        nombre_excel = (archivo.filename or "").lower()
        if not nombre_excel.endswith((".xlsx", ".xlsm", ".xls")):
            return "El archivo debe ser un Excel (.xlsx/.xlsm/.xls)", 400

        ruta_excel = UPLOADS / archivo.filename
        archivo.save(ruta_excel)

        # =============================
        # LOGO (OPCIONAL)
        # =============================
        # Si el usuario sube un logo, lo usamos en vez del logo_default.png
        logo_file = request.files.get("logo")
        ruta_logo = None
        if logo_file and getattr(logo_file, "filename", ""):
            nombre = (logo_file.filename or "").lower()
            # validación simple por extensión
            if nombre.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif")):
                # guardamos con timestamp para evitar colisiones
                ts = int(time.time())
                ext = Path(nombre).suffix
                ruta_logo = UPLOADS / f"logo_{ts}{ext}"
                logo_file.save(ruta_logo)
            else:
                return "El logo debe ser una imagen (png/jpg/jpeg/webp/bmp/tif/tiff)", 400
        # =============================
        # FIRMA (OPCIONAL)
        # =============================
        firma_file = request.files.get("firma")
        ruta_firma = None
        if firma_file and getattr(firma_file, "filename", ""):
            nombre = (firma_file.filename or "").lower()
            if nombre.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif")):
                ts = int(time.time())
                ext = Path(nombre).suffix
                ruta_firma = UPLOADS / f"firma_{ts}{ext}"
                firma_file.save(ruta_firma)
            else:
                return "La firma debe ser una imagen (png/jpg/jpeg/webp/bmp/tif/tiff)", 400

        # =============================
        # SELLO (OPCIONAL)
        # =============================
        sello_file = request.files.get("sello")
        ruta_sello = None
        if sello_file and getattr(sello_file, "filename", ""):
            nombre = (sello_file.filename or "").lower()
            if nombre.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif")):
                ts = int(time.time())
                ext = Path(nombre).suffix
                ruta_sello = UPLOADS / f"sello_{ts}{ext}"
                sello_file.save(ruta_sello)
            else:
                return "El sello debe ser una imagen (png/jpg/jpeg/webp/bmp/tif/tiff)", 400



        # =============================
        # LEER EXCEL (título + filas)
        # =============================
        titulo_llamado, texto_lote, filas = leer_items_y_descripciones_excel(ruta_excel)

        # =============================
        # MATCH (Herramientas/Materiales) - SOLO textos
        # =============================
        # Agrega a cada fila:
        #   - texto_equipos
        #   - texto_mano_obra
        #   - texto_materiales
        #   - texto_transporte
        # (No altera la parte numérica del PDF)
        filas = aplicar_match_a_filas(filas, MATCH_XLSX)

        # =============================
        # GENERACIÓN DEL PDF (NO ROMPER LO EXISTENTE)
        # =============================
        # Si no se sube logo, usamos el default.
        pdf = generar_pdf(
            filas,
            fecha,
            titulo_llamado=titulo_llamado,
            texto_lote=texto_lote,
            logo_path=(ruta_logo if ruta_logo else DEFAULT_LOGO),
            template_id=template_id,
            firma_path=ruta_firma,
            sello_path=ruta_sello,
        )

        return send_file(pdf, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)

from flask import Flask, render_template, request, send_file
from excel_utils import leer_items_y_descripciones_excel
from pdf_utils import generar_pdf
from match_utils import aplicar_match_a_filas
from report_utils import leer_reporte, parse_items_manual
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

@app.route("/", methods=["GET"])
def home():
    return render_template("home.html")


@app.route("/desglose", methods=["GET", "POST"])
def desglose():
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
        # ÍTEMS A IMPRIMIR (selector)
        # =============================
        items_mode = (request.form.get("items_mode", "all") or "all").strip().lower()
        if items_mode not in ("all", "reporte", "manual"):
            items_mode = "all"

        # Si es reporte, guardamos el archivo ahora (CSV o Excel)
        ruta_reporte = None
        if items_mode == "reporte":
            rep_file = request.files.get("reporte")
            if not rep_file or not getattr(rep_file, "filename", ""):
                return "En modo 'reporte' debes subir el archivo de reporte (.csv o .xlsx).", 400

            nombre_rep = (rep_file.filename or "").lower()
            if not nombre_rep.endswith((".csv", ".xlsx", ".xlsm", ".xls")):
                return "El reporte debe ser .csv o Excel (.xlsx/.xlsm/.xls).", 400

            ts = int(time.time())
            ext = Path(nombre_rep).suffix
            ruta_reporte = UPLOADS / f"reporte_{ts}{ext}"
            rep_file.save(ruta_reporte)

        # Si es manual, parseamos el texto
        items_manual = []
        if items_mode == "manual":
            texto_sel = request.form.get("items_manual", "")
            items_manual = parse_items_manual(texto_sel)
            if not items_manual:
                return "En modo 'selección manual' debes indicar ítems (ej: 3,11-15,18).", 400

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
        # FILTRO DE ÍTEMS SEGÚN MODO
        # =============================
        if items_mode == "manual":
            # items_manual ya viene parseado
            set_manual = set(items_manual)
            filas = [f for f in filas if int(f.get("item")) in set_manual]

        elif items_mode == "reporte":
            # Leemos el reporte y comparamos precios unitarios
            try:
                mapa_ref = leer_reporte(ruta_reporte)
            except Exception as e:
                return f"Error leyendo reporte: {e}", 400

            filas_filtradas = []
            for f in filas:
                try:
                    it = int(f.get("item"))
                except Exception:
                    continue

                oferta = f.get("precio_unitario_iva_incl")
                ref = mapa_ref.get(it)

                # Si no existe ref o ref es 0, no se puede calcular => NO se incluye (filtro estricto)
                if oferta is None or ref is None or ref == 0:
                    continue

                porcentaje = (oferta - ref) / ref * 100.0

                if porcentaje < -25.0 or porcentaje > 15.0:
                    filas_filtradas.append(f)

            filas = filas_filtradas


        # Si el filtro dejó 0 ítems, no generamos un PDF vacío
        if not filas:
            return "No hay ítems para imprimir con el criterio seleccionado.", 400

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

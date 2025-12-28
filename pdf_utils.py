import fitz  # PyMuPDF
from pathlib import Path
from costos_partes import calcular_partes_desde_cdt  # NUEVO: cálculo D/E/F/A+B desde CDT

TEMPLATE = Path("template_desglose.pdf")
OUTPUT = Path("output.pdf")

# ===============================
# CONFIGURACIÓN DE LAYOUT
# ===============================

# Columna de DESCRIPCIÓN (NO TOCAR)
X_DESCRIPCION = 270
ANCHO_DESCRIPCION = 235

# Columna de FECHA (ya ajustada en tu zip)
X_FECHA = 123

# Columna de N° ÍTEM (AJUSTABLE para probar)
# (Poné acá la X que ya encontraste)
X_ITEM = 241

Y_POSICIONES = [
    91,  # primera descripción de la hoja
    440   # segunda descripción de la hoja
]

ALTO_BLOQUE = 80

# ==========================================
# BLOQUE DE INFO (debajo de la fecha)
# ==========================================
# En el template, debajo de la fila donde va:
#   - Fecha (izquierda)
#   - Ítem (centro)
#   - Descripción (derecha)
# hay una segunda fila en blanco con 2 cajas:
#   - Izquierda: debe ir el TÍTULO DEL LLAMADO (texto de A1 o celda similar)
#   - Derecha: debe ir Unidad de medida + Presentación (por ítem)
#
# Dejamos TODO ajustable acá para que puedas moverlo rápido.

# Caja izquierda (título del llamado)
X_LLAMADO = 63                 # Ajustable: X inicio del texto del llamado
Y_LLAMADO_OFFSET = 35          # Ajustable: cuánto más abajo que 'y' se imprime
ANCHO_LLAMADO = X_DESCRIPCION + 48  # Ajustable: ancho hasta el borde antes de la columna descripción
ALTO_LLAMADO = 40              # Ajustable: alto de la caja

# Caja derecha (unidad + presentación)
X_UNI_PRES = X_DESCRIPCION + 115  # Ajustable: pequeño padding dentro de la caja derecha
Y_UNI_PRES_OFFSET = Y_LLAMADO_OFFSET         # Ajustable: mismo offset vertical que el llamado
ANCHO_UNI_PRES = round(ANCHO_DESCRIPCION/2)
ALTO_UNI_PRES = 34

# Tamaños de fuente (ajustables)
FUENTE_INFO_MAX = 7.0
FUENTE_INFO_MIN = 5.0


# ==========================================
# NUEVO: TEXTO "LOTE" (2da fila de cada tabla)
# ==========================================
# En varios Excels aparece un texto tipo "Lote ..." en la parte superior.
# Ese texto debe imprimirse en la segunda fila de cada tabla, en UNA SOLA LÍNEA.
# Si no entra, se reduce el tamaño de la letra.
#
# Coordenadas RELATIVAS al 'y' base de cada tabla (Y_POSICIONES), para que puedas
# ajustarlas fácil sin tocar el resto del código.

X_LOTE = 63                 # X inicio del texto
Y_LOTE_OFFSET = -2          # baseline relativo a 'y'
ANCHO_LOTE = 495            # ancho disponible (ajustable)

# IMPORTANTE:
# El texto de Lote debe quedar CENTRADO respecto a la TABLA (no a la hoja).
# Para eso, definimos el ancho real de la tabla (borde izquierdo y derecho).
# Por defecto lo derivamos de coordenadas ya existentes. Si querés afinarlo,
# ajustá SOLO estas dos constantes.
X_LOTE_TABLA = 59.83      # borde izquierdo real de la tabla (ajustable)
ANCHO_LOTE_TABLA = 450.79  # ancho real de la tabla (ajustable)

FUENTE_LOTE_MAX = 8.0
FUENTE_LOTE_MIN = 5.0


# ==========================================
# NUEVO: CAJAS A/B/E/F (textos desde match.xlsx)
# ==========================================
# Estas 4 cajas son SOLO texto (no afectan a los cálculos numéricos).
# Se imprimen dentro de la tabla en:
#   0) A - Equipos a Utilizar      -> texto_equipos
#   1) B - Mano de obra            -> texto_mano_obra (puede ser vacío)
#   2) E - Materiales              -> texto_materiales (puede ser vacío)
#   3) F - Transporte              -> texto_transporte
#
# IMPORTANTE:
# - "Equipos" y "Transporte" siempre se imprimen (siempre hay texto).
# - "Mano de obra" y "Materiales" pueden quedar vacíos según reglas del ítem.
#
# Coordenadas: definidas RELATIVAS al 'y' base de cada tabla (Y_POSICIONES).
# Así ajustás una sola vez y aplica a la tabla superior e inferior.

X_CAJAS_PARTES = [60, 60, 60, 60]               # X inicio de cada caja (A, B, E, F)
Y_CAJAS_PARTES_OFFSET = [81, 123, 182, 218]     # Y offset relativo a 'y'
ANCHO_CAJAS_PARTES = [160, 160, 160, 160]       # Ancho por caja
# IMPORTANTE:
# En el template real, las cajas B (Mano de obra) y F (Transporte) tienen más altura.
# Si el alto es demasiado chico, PyMuPDF puede NO insertar el texto (insert_textbox retorna < 0)
# y el resultado es que no se ve nada en el PDF.
#
# Por eso dejamos un alto más generoso por defecto para B y F.
ALTO_CAJAS_PARTES = [16, 20, 17, 20]            # Alto por caja

# Padding interno (por si querés que el texto no quede pegado al borde)
PAD_X_PARTES = 2
PAD_Y_PARTES = 0

# ===============================
# NUEVO: NÚMEROS DEL RESUMEN (CDT .. CU+IVA)
# ===============================
# A partir del monto TOTAL (IVA incluido) del Excel, calculamos los valores desde
# "Costo Directo Total" hasta "Costo Unitario Adoptado".
#
# REGLAS IMPORTANTES:
# - Nunca imprimimos decimales (todo entero).
# - Nunca imprimimos negativos (clamp >= 0).
#
# Porcentajes base (se pueden ajustar fácilmente):
PCT_CDT = 0.75  # CDT = CU(sin IVA) * 75%
PCT_GG = 0.14   # GG teórico = CU(sin IVA) * 14%  (luego se ajusta para cuadrar)
PCT_BEL = 0.10  # Bel teórico = CU(sin IVA) * 10%

# Coordenadas (vectores) de los números del resumen en el PDF.
# Son ABSOLUTAS en la página y existen para la tabla de ARRIBA (index 0)
# y para la tabla de ABAJO (index 1).
#
# Cada fila corresponde a:
#   0: CDT
#   1: GG
#   2: Bel (en el template aparece como "Impuestos y retenciones")
#   3: CU
#   4: IVA
#   5: CU + IVA
RESUMEN_FILAS = ["CDT", "GG", "BEL", "CU", "IVA", "CU_IVA"]

# Columna donde se imprimen los valores (x0, x1). Si querés mover la columna completa,
# ajustá SOLO estos dos números.
X_RESUMEN_RIGHT = 508.00  # Ajustable: borde derecho (dentro de la celda) para alinear a la derecha

# Si querés mover SOLO alguna fila, podés cambiar el valor por fila (vector).
X_RESUMEN_RIGHTS = [X_RESUMEN_RIGHT] * len(RESUMEN_FILAS)

# Y (vectores) por tabla:
# Un valor por fila de RESUMEN_FILAS (posición baseline para insert_text).
Y_RESUMEN_BASE = [
    # Tabla de ARRIBA (index 0)
    [334.5, 344, 354.68, 374.36, 383.84, 393.44],
    # Tabla de ABAJO (index 1)
    [682.5, 692, 702.68, 722.48, 731.96, 741.56],
]



# ===============================
# NUEVO: NÚMEROS DE PARTES (A+B, D, E, F)
# ===============================
# A partir del CDT (ya calculado en el resumen), repartimos el CDT en:
#   - D (Ejecución) = A + B  (C = 1)
#   - E (Materiales)
#   - F (Transporte)
#
# IMPORTANTE:
# - Estos números NO afectan el resumen (CDT..CU+IVA). Solo se imprimen en el cuerpo.
# - Todo es entero, sin decimales y sin negativos.
#
# Coordenadas (vectores) ABSOLUTAS en la página:
# - Tabla superior (index 0)
# - Tabla inferior (index 1)
#
# Orden de filas:
#   0: (A+B)   -> celda "Costo de Producción (A+B)"
#   1: D       -> celda de "D) - Costo Unitario de la Ejecución ..."
#   2: E       -> fila "E) TOTAL Gs."
#   3: F       -> fila "F) TOTAL Gs."
PARTES_NUM_KEYS = ["AB", "D", "E", "F"]


# ----------------------------------------------------------
# NUEVO: TOTALES A y B (filas propias en el template)
# ----------------------------------------------------------
# En el PDF, A) (Equipos) y B) (Mano de Obra) tienen sus propias filas "TOTAL Gs.".
# Estos totales se deducen desde D porque:
#   - Con C = 1, D = A + B
#   - Para ítems de materiales: B = 0  => A = D
#   - Para ítems NO materiales: A = 10% de D, B = D - A
#
# IMPORTANTE:
# - No tocamos coordenadas existentes (D/E/F/AB). Estas son nuevas y editables.
# - Si querés ajustar posiciones, tocá SOLO Y_AB_TOTALES_BASE_TOP o X_AB_TOTALES_RIGHT.

AB_TOTALES_KEYS = ["A", "B"]

# Misma columna derecha que el resto de totales (ajustable si hace falta).
X_AB_TOTALES_RIGHT = 508.00
X_AB_TOTALES_RIGHTS = [X_AB_TOTALES_RIGHT] * len(AB_TOTALES_KEYS)

# Baselines (tabla de ARRIBA) para:
#   0: A) TOTAL Gs.
#   1: B) TOTAL Gs.
# Estos valores están alineados al template (y1 de las etiquetas).
Y_AB_TOTALES_BASE_TOP = [195.16, 236.32]

# Y por tabla (arriba / abajo)
Y_AB_TOTALES_BASE = [
    Y_AB_TOTALES_BASE_TOP,
    [y + 348.12 for y in Y_AB_TOTALES_BASE_TOP],
]

# Columna derecha donde se alinean los números (mismo criterio que el resumen).
X_PARTES_RIGHT = 508.00
X_PARTES_RIGHTS = [X_PARTES_RIGHT] * len(PARTES_NUM_KEYS)

# Y base (baseline) para la tabla de ARRIBA.
# Si querés ajustar rápido, editá SOLO este vector y/o DELTA_TABLAS.
Y_PARTES_BASE_TOP = [250.00, 263.00, 299.00, 326.00]

# Diferencia vertical entre tabla de arriba y tabla de abajo (se ve en el template).
# Nota: en el mismo archivo ya se usa un delta muy similar para el resumen.
DELTA_TABLAS = 348.12

# Y por tabla (arriba / abajo)
Y_PARTES_BASE = [
    Y_PARTES_BASE_TOP,
    [y + DELTA_TABLAS for y in Y_PARTES_BASE_TOP],
]


# ===============================
# NUEVO: DETALLES DE A (EQUIPOS) Y F (TRANSPORTE)
# ===============================
# En el template hay filas de detalle para:
#   - A) Equipos: Horas / Costo por hora / Costo total horario
#   - F) Transporte: DTM / Consumo / Costo unitario / Costo total unitario
#
# IMPORTANTE (según tu última indicación):
# - En la fila ANTERIOR a "A) TOTAL Gs." se debe repetir el MISMO total A en la columna
#   "Costo Total Horario Gs.".
# - En la fila ANTERIOR a "F) TOTAL Gs." se debe repetir el MISMO total F en la columna
#   "Costo Total Unitario Gs.".
#
# Para que sea fácil ajustar, dejamos coordenadas en VECTORES.
# Las Y son baselines (para insert_text) y existen para tabla ARRIBA (0) y ABAJO (1).

# ---- A) EQUIPOS (detalle) ----
# Orden de columnas a imprimir:
#   0: Horas de C/Equipos          (entero)
#   1: Costos Horario Gs.          (entero)
#   2: Costo Total Horario Gs.     (entero)  <-- se repite el TOTAL A
A_DET_COLS = ["HORAS", "COSTO_HORA", "COSTO_TOTAL"]

# X (alineación derecha dentro de cada celda)
# Ajustá estos X si querés mover los números dentro de sus columnas.
X_A_DET_RIGHTS = [291.0, 382.0, 508.0]

# Y base (baseline) por tabla (arriba/abajo)
Y_A_DET_BASE_TOP = 183.0
Y_A_DET_BASE = [
    Y_A_DET_BASE_TOP,
    Y_A_DET_BASE_TOP + DELTA_TABLAS,
]

# ---- F) TRANSPORTE (detalle) ----
# Orden de columnas a imprimir:
#   0: DTM (Km.)                   (float con 2 decimales y coma)
#   1: Consumo                     (float con 2 decimales y coma)
#   2: Costo Unitario Gs           (entero)
#   3: Costo Total Unitario Gs     (entero)  <-- se repite el TOTAL F
F_DET_COLS = ["DTM", "CONSUMO", "COSTO_UNIT", "COSTO_TOTAL"]

# X (alineación derecha dentro de cada celda)
X_F_DET_RIGHTS = [259.0, 297.0, 382.0, 508.0]

# Y base (baseline) por tabla (arriba/abajo)
Y_F_DET_BASE_TOP = 316.5
Y_F_DET_BASE = [
    Y_F_DET_BASE_TOP,
    Y_F_DET_BASE_TOP + DELTA_TABLAS,
]

# Parámetros de la generación del detalle (editables)
DET_A_HORAS_MIN = 2
DET_A_HORAS_MAX = 10

DET_F_COSTO_UNIT_MIN = 8000
DET_F_COSTO_UNIT_MAX = 11000
DET_F_CONSUMO = 0.05


# ===============================
# NUEVO: DETALLES DE B (MANO DE OBRA) Y E (MATERIALES)
# ===============================
# En el template hay filas de detalle para:
#   - B) Mano de obra: Cantidad / Horas / Costo / Costo total
#   - E) Materiales: Consumo / Costo unitario / Costo total unitario
#
# Reglas:
# - Estos detalles SOLO se imprimen si corresponde:
#     * B se imprime si total_B > 0
#     * E se imprime si total_E > 0
# - "Cantidad" se toma del Excel (columna cantidad).
# - B: Horas aleatorio entero entre 1 y 3.
# - B: Costo = total_B / (cantidad * horas).
# - E: Consumo = cantidad (mismo valor).
# - E: Costo unitario = total_E / consumo.
#
# Importante: solo enteros; sin negativos; sin coma en números (excepto DTM/Consumo de Transporte).

# ---- B) MANO DE OBRA (detalle) ----
# Orden de columnas a imprimir:
#   0: Cantidad                     (entero)
#   1: Horas                        (entero)
#   2: Costo (Gs)                   (entero)
#   3: Costo Total (Gs)             (entero)  <-- se repite el TOTAL B
B_DET_COLS = ["CANTIDAD", "HORAS", "COSTO", "COSTO_TOTAL"]

# X (alineación derecha dentro de cada celda)
X_B_DET_RIGHTS = [259.0, 297.0, 382.0, 508.0]

# Y base (baseline) por tabla (arriba/abajo)
Y_B_DET_BASE_TOP = 224.0
Y_B_DET_BASE = [
    Y_B_DET_BASE_TOP,
    Y_B_DET_BASE_TOP + DELTA_TABLAS,
]

# Parámetros del detalle B
DET_B_HORAS_MIN = 1
DET_B_HORAS_MAX = 3

# ---- E) MATERIALES (detalle) ----
# Orden de columnas a imprimir:
#   0: Consumo                      (entero)
#   1: Costo Unitario (Gs)          (entero)
#   2: Costo Total Unitario (Gs)    (entero)  <-- se repite el TOTAL E
E_DET_COLS = ["CONSUMO", "COSTO_UNIT", "COSTO_TOTAL"]

# X (alineación derecha dentro de cada celda)
X_E_DET_RIGHTS = [297.0, 382.0, 508.0]

# Y base (baseline) por tabla (arriba/abajo)
Y_E_DET_BASE_TOP = 286.0
Y_E_DET_BASE = [
    Y_E_DET_BASE_TOP,
    Y_E_DET_BASE_TOP + DELTA_TABLAS,
]

# ===============================
# OPCIÓN B: TAPAR SEGUNDA TABLA
# ===============================
TAPAR_SEGUNDA_TABLA_EN_ULTIMA_HOJA_SI_IMPAR = True

# Ajustable: desde qué Y hacia abajo se tapa (para ocultar la tabla inferior).
# Ponemos un valor razonable cerca del inicio de la tabla inferior.
# Si tapa de más o de menos, ajustá SOLO esta constante.
TAPAR_Y0 = 420


# ==========================================
# NUEVO: LOGO (esquina superior derecha)
# ==========================================
# Se inserta 1 vez por página (cuando se copia la página del template).
# Si el usuario no sube un logo, se usa logo_default.png.
# Coordenadas ajustables por constantes:

LOGO_W = 75                 # ancho del logo (ajustable)
LOGO_H = 32                 # alto del logo (ajustable)
LOGO_SCALE = 1.10          # agrandar logo 10% (mantiene esquina superior derecha)
LOGO_MARGIN_RIGHT = 18      # margen derecho (ajustable)
LOGO_MARGIN_TOP = 14        # margen superior (ajustable)


def insertar_texto_autoajustado(page, rect, texto):
    """
    Inserta texto respetando:
    - máximo 3 líneas
    - tamaño máximo 7.0
    - tamaño mínimo 4.5

    (Esta función la usabas para la descripción: NO TOCAR la lógica.)
    """
    MAX_LINEAS = 3
    FONT_MAX = 7.0
    FONT_MIN = 4.5
    PASO = 0.1

    texto = str(texto).strip()

    size = FONT_MAX
    while size >= FONT_MIN:
        text_length = fitz.get_text_length(
            texto,
            fontname="helv",
            fontsize=size
        )

        alto_linea = size * 1.2
        lineas_estimadas = (text_length / rect.width) * 1.15
        alto_maximo = MAX_LINEAS * alto_linea

        if lineas_estimadas * alto_linea <= alto_maximo:
            page.insert_textbox(
                rect,
                texto,
                fontsize=size,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT
            )
            return size

        size -= PASO

    page.insert_textbox(
        rect,
        texto,
        fontsize=FONT_MIN,
        fontname="helv",
        color=(0, 0, 0),
        align=fitz.TEXT_ALIGN_LEFT
    )

    return FONT_MIN


def insertar_info_autoajustada(page, rect, texto, max_lineas=2):
    """
    Inserta el bloque de info debajo de la fecha (título / unidad+presentación)
    con autoajuste de fuente para que no se salga de la caja.

    - max_lineas: por defecto 2 (porque pediste 2 líneas: unidad y presentación)
    - fuente max/min: configurables arriba (FUENTE_INFO_MAX / FUENTE_INFO_MIN)
    """
    texto = str(texto).strip()
    size = FUENTE_INFO_MAX
    paso = 0.1

    while size >= FUENTE_INFO_MIN:
        # estimación simple: largo en puntos / ancho -> lineas
        text_length = fitz.get_text_length(texto, fontname="helv", fontsize=size)
        alto_linea = size * 1.2
        lineas_estimadas = (text_length / rect.width) * 1.15
        alto_maximo = max_lineas * alto_linea

        if lineas_estimadas * alto_linea <= alto_maximo:
            page.insert_textbox(
                rect,
                texto,
                fontsize=size,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT
            )
            return size

        size -= paso

    page.insert_textbox(
        rect,
        texto,
        fontsize=FUENTE_INFO_MIN,
        fontname="helv",
        color=(0, 0, 0),
        align=fitz.TEXT_ALIGN_LEFT
    )

    return FUENTE_INFO_MIN





def insertar_texto_partes_autoajustado(page, rect, texto):
    """
    NUEVO:
    Inserta texto en cajas chicas (Equipos / Mano de obra / Materiales / Transporte)
    usando autoajuste similar a la descripción, pero con fuentes más pequeñas.

    - Máximo 3 líneas (por defecto) para que entren frases como
      "Supervisor, técnicos oficiales y técnicos ayudantes".
    - NO toca la lógica de la descripción ni la de los números.
    """
    MAX_LINEAS = 3
    FONT_MAX = 6.0
    FONT_MIN = 3.0
    PASO = 0.1

    texto = str(texto).strip()

    size = FONT_MAX
    while size >= FONT_MIN:
        text_length = fitz.get_text_length(texto, fontname="helv", fontsize=size)
        alto_linea = size * 1.2
        lineas_estimadas = (text_length / rect.width) * 1.15
        alto_maximo = MAX_LINEAS * alto_linea

        if lineas_estimadas * alto_linea <= alto_maximo:
            # insert_textbox puede FALLAR (retorna < 0) si la caja es muy chica.
            # En ese caso seguimos bajando el tamaño.
            ret = page.insert_textbox(
                rect,
                texto,
                fontsize=size,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT
            )
            if ret >= 0:
                return size

        size -= PASO

    # Último intento con FONT_MIN
    ret = page.insert_textbox(
        rect,
        texto,
        fontsize=FONT_MIN,
        fontname="helv",
        color=(0, 0, 0),
        align=fitz.TEXT_ALIGN_LEFT
    )

    # Si todavía falla, usamos insert_text (no depende del alto de la caja).
    if ret < 0:
        page.insert_text(
            (rect.x0, rect.y0 + FONT_MIN),
            texto,
            fontsize=FONT_MIN,
            fontname="helv",
            color=(0, 0, 0)
        )

    return FONT_MIN


def insertar_texto_una_linea_autofit(page, x0, y_baseline, ancho, texto, font_max, font_min, centrado=True):
    """Inserta *UNA SOLA LÍNEA* reduciendo la fuente si hace falta.

    Se usa para el texto de "Lote" porque:
    - No puede partirse en 2 líneas
    - Debe entrar en el ancho disponible
    - Debe quedar centrado dentro del ancho de la tabla (centrado=True)

    NOTA: solo el texto de Lote usa esta función actualmente.
    """
    texto = " ".join(str(texto).replace("\n", " ").split()).strip()
    paso = 0.1

    size = font_max
    while size >= font_min:
        w = fitz.get_text_length(texto, fontname="helv", fontsize=size)
        if w <= ancho:
            x = x0 + (ancho - w) / 2 if centrado else x0
            page.insert_text(
                (x, y_baseline),
                texto,
                fontsize=size,
                fontname="helv",
                color=(0, 0, 0)
            )
            return size
        size -= paso

    # Último intento con font_min
    w = fitz.get_text_length(texto, fontname="helv", fontsize=font_min)
    x = x0 + (ancho - w) / 2 if centrado and w <= ancho else x0
    page.insert_text(
        (x, y_baseline),
        texto,
        fontsize=font_min,
        fontname="helv",
        color=(0, 0, 0)
    )
    return font_min
def insertar_lote(page, y_base, texto_lote):
    """Imprime el texto de Lote en la 2da fila de cada tabla (una línea)."""
    if not texto_lote:
        return

    insertar_texto_una_linea_autofit(
        page,
        X_LOTE_TABLA,
        y_base + Y_LOTE_OFFSET,
        ANCHO_LOTE_TABLA,
        texto_lote,
        FUENTE_LOTE_MAX,
        FUENTE_LOTE_MIN,
    )


def insertar_logo_en_pagina(page, logo_path):
    """Inserta el logo en la esquina superior derecha de la página.

    IMPORTANTE:
    - Mantiene SIEMPRE las proporciones originales del logo (sin distorsión).
    - LOGO_W/LOGO_H representan un *bounding box* máximo. LOGO_SCALE permite agrandarlo.
    - El logo se ancla por la esquina superior derecha (márgenes constantes).
    """
    if not logo_path:
        return

    try:
        p = Path(logo_path)
        if not p.exists():
            return

        # Bounding box máximo (con escala)
        max_w = LOGO_W * LOGO_SCALE
        max_h = LOGO_H * LOGO_SCALE

        # Tamaño real de la imagen (para mantener proporciones)
        pix = fitz.Pixmap(str(p))
        img_w, img_h = float(pix.width), float(pix.height)
        pix = None  # liberar

        if img_w <= 0 or img_h <= 0:
            return

        ratio = img_w / img_h

        # Ajustar (fit) dentro del bounding box manteniendo aspecto
        if (max_w / max_h) >= ratio:
            # limita por alto
            h = max_h
            w = h * ratio
        else:
            # limita por ancho
            w = max_w
            h = w / ratio

        # Anclar arriba-derecha
        x1 = page.rect.width - LOGO_MARGIN_RIGHT
        y0 = LOGO_MARGIN_TOP
        x0 = x1 - w
        y1 = y0 + h

        page.insert_image(
            fitz.Rect(x0, y0, x1, y1),
            filename=str(p),
        )
    except Exception:
        # Si el logo no se puede insertar por algún motivo, no rompemos la generación.
        return


def _round_half_up(x):
    """
    Emula el redondeo típico de Excel (ROUND) para 0 decimales,
    evitando el "bankers rounding" de Python.

    Retorna int.
    """
    from decimal import Decimal, ROUND_HALF_UP
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _clamp_nonneg_int(x):
    """
    Asegura entero >= 0.
    """
    try:
        xi = int(x)
    except Exception:
        xi = 0
    return xi if xi >= 0 else 0


def _format_gs(n):
    """
    Formato Guaraníes:
    - entero
    - separador de miles con punto (.)
    - sin decimales, sin coma
    """
    n = _clamp_nonneg_int(n)
    return f"{n:,}".replace(",", ".")


def _format_float_coma(x, decimales=2):
    """
    Formato numérico con:
    - separador de miles con punto (.)
    - separador decimal con coma (,)
    - decimales fijos

    IMPORTANTE:
    Solo usamos esto para DTM y Consumo (como pediste).
    """
    try:
        xf = float(x)
    except Exception:
        xf = 0.0

    # Clamp a no-negativo
    if xf < 0:
        xf = 0.0

    # Formato estilo "1,234.56" y luego convertir a "1.234,56"
    s = f"{xf:,.{decimales}f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def _insertar_texto_derecha(page, x_right, y, texto, fontsize=8):
    """
    Inserta texto alineado a la derecha, usando insert_text (robusto y consistente).
    """
    texto = str(texto)
    ancho = fitz.get_text_length(texto, fontname="helv", fontsize=fontsize)
    x = x_right - ancho
    page.insert_text((x, y), texto, fontsize=fontsize, fontname="helv", color=(0, 0, 0))


def _insertar_detalles_a_y_f(page, tabla_index, seed_int, partes):
    """
    Inserta los detalles de:
      - A) Equipos: horas / costo por hora / costo total (repite total A)
      - F) Transporte: dtm / consumo / costo unit / costo total (repite total F)

    No modifica ni interfiere con:
      - Resumen (CDT..CU+IVA)
      - Partes (A+B, D, E, F) ni Totales A/B ya impresos
      - Textos de match (equipos/mano obra/materiales/transporte)
    """
    import random

    rng = random.Random(int(seed_int) if seed_int is not None else 0)

    total_a = _clamp_nonneg_int(partes.get("A", 0))
    total_f = _clamp_nonneg_int(partes.get("F", 0))

    # -------------------------
    # A) Equipos (detalle)
    # -------------------------
    # Elegimos horas entre 2 y 10 (inclusive). Para que quede prolijo, si existe un divisor
    # en ese rango lo preferimos (así total_a / horas queda exacto).
    horas_candidates = [h for h in range(DET_A_HORAS_MIN, DET_A_HORAS_MAX + 1) if h > 0 and total_a % h == 0]
    if horas_candidates:
        horas = rng.choice(horas_candidates)
    else:
        horas = rng.randint(DET_A_HORAS_MIN, DET_A_HORAS_MAX)

    horas = _clamp_nonneg_int(horas) if horas > 0 else 1

    # Costo por hora (entero) derivado del total
    costo_hora = 0
    if horas > 0:
        costo_hora = _clamp_nonneg_int(total_a // horas)

    # Repetimos total A en la columna "Costo Total Horario Gs."
    y_a = Y_A_DET_BASE[tabla_index]
    _insertar_texto_derecha(page, X_A_DET_RIGHTS[0], y_a, _format_gs(horas), fontsize=8)
    _insertar_texto_derecha(page, X_A_DET_RIGHTS[1], y_a, _format_gs(costo_hora), fontsize=8)
    _insertar_texto_derecha(page, X_A_DET_RIGHTS[2], y_a, _format_gs(total_a), fontsize=8)

    # -------------------------
    # F) Transporte (detalle)
    # -------------------------
    consumo = float(DET_F_CONSUMO)
    costo_unit = rng.randint(DET_F_COSTO_UNIT_MIN, DET_F_COSTO_UNIT_MAX)
    costo_unit = _clamp_nonneg_int(costo_unit) if costo_unit > 0 else DET_F_COSTO_UNIT_MIN

    denom = consumo * float(costo_unit)
    dtm = 0.0
    if denom > 0:
        dtm = float(total_f) / denom

    y_f = Y_F_DET_BASE[tabla_index]
    _insertar_texto_derecha(page, X_F_DET_RIGHTS[0], y_f, _format_float_coma(dtm, 2), fontsize=8)
    _insertar_texto_derecha(page, X_F_DET_RIGHTS[1], y_f, _format_float_coma(consumo, 2), fontsize=8)
    _insertar_texto_derecha(page, X_F_DET_RIGHTS[2], y_f, _format_gs(costo_unit), fontsize=8)
    _insertar_texto_derecha(page, X_F_DET_RIGHTS[3], y_f, _format_gs(total_f), fontsize=8)

def _insertar_detalles_b_y_e(page, tabla_index, seed_int, partes, cantidad_excel):
    """Inserta los detalles de:

    - B) Mano de obra (solo si total_B > 0)
        * Cantidad (del Excel)
        * Horas por trabajador (aleatorio entero 1..3)
        * Costo (Gs) = total_B / (cantidad * horas)
        * Costo total (Gs) = total_B (se repite)

    - E) Materiales (solo si total_E > 0)
        * Consumo (= cantidad del Excel)
        * Costo unitario (Gs) = total_E / consumo
        * Costo total unitario (Gs) = total_E (se repite)

    Importante:
    - No imprime nada si la parte no corresponde (fila vacía en el PDF).
    - No usa decimales (todo entero).
    """
    import random

    rng = random.Random((int(seed_int) if seed_int is not None else 0) + 1337)

    total_b = _clamp_nonneg_int(partes.get("B", 0))
    total_e = _clamp_nonneg_int(partes.get("E", 0))

    # Cantidad (Excel). La usamos para Mano de obra y también como Consumo en Materiales.
    try:
        qty = _round_half_up(cantidad_excel) if cantidad_excel is not None else 0
    except Exception:
        qty = 0
    qty = _clamp_nonneg_int(qty)
    if qty <= 0:
        qty = 1

    # -------------------------
    # B) Mano de obra (detalle)
    # -------------------------
    if total_b > 0:
        # Preferimos horas que den división exacta si existen
        horas_candidates = [
            h for h in range(DET_B_HORAS_MIN, DET_B_HORAS_MAX + 1)
            if h > 0 and (qty * h) > 0 and (total_b % (qty * h) == 0)
        ]
        if horas_candidates:
            horas = rng.choice(horas_candidates)
        else:
            horas = rng.randint(DET_B_HORAS_MIN, DET_B_HORAS_MAX)

        horas = _clamp_nonneg_int(horas) if horas > 0 else 1
        denom = qty * horas
        if denom <= 0:
            denom = 1

        costo = _clamp_nonneg_int(_round_half_up(total_b / denom))
        if costo <= 0 and total_b > 0:
            costo = 1

        y_b = Y_B_DET_BASE[tabla_index]
        _insertar_texto_derecha(page, X_B_DET_RIGHTS[0], y_b, str(qty), fontsize=8)
        _insertar_texto_derecha(page, X_B_DET_RIGHTS[1], y_b, str(horas), fontsize=8)
        _insertar_texto_derecha(page, X_B_DET_RIGHTS[2], y_b, _format_gs(costo), fontsize=8)
        _insertar_texto_derecha(page, X_B_DET_RIGHTS[3], y_b, _format_gs(total_b), fontsize=8)

    # -------------------------
    # E) Materiales (detalle)
    # -------------------------
    if total_e > 0:
        consumo = qty if qty > 0 else 1
        costo_unit = _clamp_nonneg_int(_round_half_up(total_e / consumo)) if consumo > 0 else total_e
        if costo_unit <= 0 and total_e > 0:
            costo_unit = 1

        y_e = Y_E_DET_BASE[tabla_index]
        _insertar_texto_derecha(page, X_E_DET_RIGHTS[0], y_e, str(consumo), fontsize=8)
        _insertar_texto_derecha(page, X_E_DET_RIGHTS[1], y_e, _format_gs(costo_unit), fontsize=8)
        _insertar_texto_derecha(page, X_E_DET_RIGHTS[2], y_e, _format_gs(total_e), fontsize=8)

def _calcular_resumen_desde_total(total_iva_incl):
    """
    Construye los valores del resumen desde el TOTAL (IVA incluido).

    Devuelve dict:
        {
          "CDT": ...,
          "GG": ...,
          "BEL": ...,
          "CU": ...,
          "IVA": ...,
          "CU_IVA": ...
        }

    Nota:
    - No se imprimen decimales.
    - No se permiten negativos.
    - GG se ajusta (Opción A) para que CU = CDT + GG + BEL, donde CU = (CU+IVA) - IVA.
    """
    if total_iva_incl is None:
        total_iva_incl = 0

    # Total IVA incluido (entero)
    cu_iva = _round_half_up(total_iva_incl)

    # IVA y CU (sin IVA)
    iva = _round_half_up(cu_iva / 11)
    cu = cu_iva - iva

    # Evitar negativos por cualquier motivo (aunque no debería)
    iva = _clamp_nonneg_int(iva)
    cu = _clamp_nonneg_int(cu)

    # CDT y Bel teórico (sobre CU sin IVA)
    cdt = _round_half_up(cu * PCT_CDT)
    bel = _round_half_up(cu * PCT_BEL)

    # Clamps para evitar pasar CU por redondeos
    cdt = min(_clamp_nonneg_int(cdt), cu)
    bel = min(_clamp_nonneg_int(bel), cu - cdt)

    # Opción A (simplificada): GG absorbe el ajuste para que cierre
    gg = cu - cdt - bel
    gg = _clamp_nonneg_int(gg)

    return {
        "CDT": cdt,
        "GG": gg,
        "BEL": bel,          # en el template figura como "Impuestos y retenciones"
        "CU": cdt + gg + bel,
        "IVA": iva,
        "CU_IVA": cu_iva,
    }


def _insertar_numero_resumen(page, tabla_index, fila_index, valor):
    """
    Inserta un número del resumen en el PDF usando los vectores:
      - X_RESUMEN_RIGHTS
      - Y_RESUMEN_BASE

    Usamos insert_text (en vez de insert_textbox) para evitar overflow de la celda y
    que se pisen los valores entre filas.
    """
    texto = _format_gs(valor)

    x_right = X_RESUMEN_RIGHTS[fila_index]
    y = Y_RESUMEN_BASE[tabla_index][fila_index]

    # Alineación a la derecha: calculamos ancho y ubicamos el inicio.
    ancho = fitz.get_text_length(texto, fontname="helv", fontsize=8)
    x = x_right - ancho

    page.insert_text(
        (x, y),
        texto,
        fontsize=8,
        fontname="helv",
        color=(0, 0, 0)
    )




def _insertar_numero_partes(page, tabla_index, fila_index, valor):
    """
    Inserta un número de partes (A+B, D, E, F) usando los vectores:
      - X_PARTES_RIGHTS
      - Y_PARTES_BASE

    Mismo criterio que _insertar_numero_resumen: alineado a la derecha con insert_text.
    """
    texto = _format_gs(valor)

    x_right = X_PARTES_RIGHTS[fila_index]
    y = Y_PARTES_BASE[tabla_index][fila_index]

    ancho = fitz.get_text_length(texto, fontname="helv", fontsize=8)
    x = x_right - ancho

    page.insert_text(
        (x, y),
        texto,
        fontsize=8,
        fontname="helv",
        color=(0, 0, 0)
    )


def _insertar_numero_ab_totales(page, tabla_index, fila_index, valor):
    """
    Inserta los totales A) y B) en sus filas propias (A) TOTAL Gs. / B) TOTAL Gs.)
    usando los vectores:
      - X_AB_TOTALES_RIGHTS
      - Y_AB_TOTALES_BASE
    """
    texto = _format_gs(valor)

    x_right = X_AB_TOTALES_RIGHTS[fila_index]
    y = Y_AB_TOTALES_BASE[tabla_index][fila_index]

    ancho = fitz.get_text_length(texto, fontname="helv", fontsize=8)
    x = x_right - ancho

    page.insert_text(
        (x, y),
        texto,
        fontsize=8,
        fontname="helv",
        color=(0, 0, 0)
    )

def generar_pdf(filas, fecha, titulo_llamado="", texto_lote="", logo_path=None):
    """
    Genera un PDF usando el template existente.
    Coloca 2 ítems por hoja.

    - Solo imprimimos filas válidas (excel_utils filtra por item).
    - Si la cantidad total es impar, la última hoja debe mostrar SOLO una tabla:
      => Opción B: tapar visualmente la tabla inferior.

    Bloque debajo de fecha:
    - Debajo de la fecha (y en el bloque superior de cada tabla), imprimimos:
        * Izquierda: título del llamado (desde el Excel, fila 1)
        * Derecha: 'Unidad de medida' y 'Presentación' (por cada ítem)
      Las coordenadas quedan arriba como constantes para que puedas ajustarlas fácil.

    NUEVO:
    - Imprime los valores del resumen desde CDT hasta CU+IVA en ambas tablas.
      (Se calcula desde el TOTAL IVA incluido del Excel).

    NUEVO (encabezado):
    - Texto de "Lote" (una sola línea) en la 2da fila de cada tabla.
    - Logo en esquina superior derecha de cada página (default o subido por usuario).
    """
    template_doc = fitz.open(TEMPLATE)
    doc = fitz.open()

    total = len(filas)

    for i, fila in enumerate(filas):

        # Esperamos dict:
        #   {"item": ..., "descripcion": "...", "unidad_medida": "...", "presentacion": "...",
        #    "cantidad": ..., "precio_unitario_iva_incl": ..., "precio_total_iva_incl": ...}
        if isinstance(fila, dict):
            texto = fila.get("descripcion", "")
            numero_item = fila.get("item", "")
            unidad_medida = fila.get("unidad_medida", "")
            presentacion = fila.get("presentacion", "")
        else:
            # fallback por si llega algo inesperado
            texto = str(fila)
            numero_item = ""
            unidad_medida = ""
            presentacion = ""

        posicion_en_hoja = i % 2

        if posicion_en_hoja == 0:
            doc.insert_pdf(template_doc, from_page=0, to_page=0)
            page = doc[-1]
            # Logo: se inserta UNA VEZ por página (al crearla)
            insertar_logo_en_pagina(page, logo_path)
        else:
            page = doc[-1]

        y = Y_POSICIONES[posicion_en_hoja]

        # -------------------------------
        # FECHA (como ya te funciona)
        # -------------------------------
        AJUSTE_BASELINE = 17

        page.insert_text(
            (X_FECHA, y + AJUSTE_BASELINE),
            fecha,
            fontsize=8,
            fontname="helv",
            color=(0, 0, 0)
        )

        # -------------------------------
        # N° ÍTEM
        # - Misma Y que la fecha (y + AJUSTE_BASELINE)
        # - Ceros a la izquierda si es numérico: 1->001, 12->012, 101->101
        # -------------------------------
        item_txt = str(numero_item).strip()
        if item_txt.isdigit():
            item_txt = item_txt.zfill(3)

        page.insert_text(
            (X_ITEM, y + AJUSTE_BASELINE),
            item_txt,
            fontsize=8,
            fontname="helv",
            color=(0, 0, 0)
        )

        # ---------------------------------------------
        # NUEVO: TEXTO "LOTE" (una sola línea)
        # ---------------------------------------------
        # Se imprime en la 2da fila de cada tabla. Si no entra, se reduce la fuente.
        insertar_lote(page, y, texto_lote)

        # ---------------------------------------------
        # NUEVO: NÚMEROS DEL RESUMEN (CDT..CU+IVA)
        # ---------------------------------------------
        # El monto base que usamos es el TOTAL (IVA incluido): "Precio total".
        # Si por alguna razón no viniera, intentamos reconstruirlo con:
        #   (precio_unitario_iva_incl * cantidad)
        total_iva_incl = None
        if isinstance(fila, dict):
            total_iva_incl = fila.get("precio_total_iva_incl", None)
            if total_iva_incl is None:
                pu = fila.get("precio_unitario_iva_incl", None)
                qty = fila.get("cantidad", None)
                if pu is not None and qty is not None:
                    total_iva_incl = pu * qty

        resumen = _calcular_resumen_desde_total(total_iva_incl)

        # Posición en hoja: 0 = tabla de arriba, 1 = tabla de abajo
        tabla_index = posicion_en_hoja

        # Insertar todos los valores del resumen usando vectores de coordenadas
        for idx_fila, key in enumerate(RESUMEN_FILAS):
            _insertar_numero_resumen(page, tabla_index, idx_fila, resumen.get(key, 0))

        # ---------------------------------------------
        # NUEVO: D / E / F (y A+B) a partir del CDT
        # ---------------------------------------------
        # NOTA IMPORTANTE:
        # - No tocamos la lógica del resumen (ya funciona).
        # - Solo usamos el CDT ya calculado y el tipo de ítem (match_utils) para repartir.
        # - Herramientas / Transporte son textos (se imprimen siempre) y NO dependen de estos números.
        tipo_item = "ambiguo"
        if isinstance(fila, dict):
            tipo_item = fila.get("tipo_item", "ambiguo")

        partes = calcular_partes_desde_cdt(resumen.get("CDT", 0), tipo_item)

        # Insertamos: (A+B), D, E, F (en ambas tablas)
        valores_partes = [
            partes.get("AB", 0),
            partes.get("D", 0),
            partes.get("E", 0),
            partes.get("F", 0),
        ]
        for idx_parte, val in enumerate(valores_partes):
            _insertar_numero_partes(page, tabla_index, idx_parte, val)

        # NUEVO: Totales A) y B) (filas propias)
        valores_ab = [
            partes.get("A", 0),
            partes.get("B", 0),
        ]
        for idx_ab, val in enumerate(valores_ab):
            _insertar_numero_ab_totales(page, tabla_index, idx_ab, val)

        # ---------------------------------------------
        # NUEVO: DETALLE DE A (Equipos) Y F (Transporte)
        # ---------------------------------------------
        # Usamos un seed estable por ítem para que los aleatorios sean REPRODUCIBLES.
        if item_txt.isdigit():
            seed_int = int(item_txt)
        else:
            seed_int = abs(hash(str(texto))) % 1000000

        _insertar_detalles_a_y_f(page, tabla_index, seed_int, partes)

        # ---------------------------------------------
        # NUEVO: DETALLE DE B (Mano de obra) Y E (Materiales)
        # ---------------------------------------------
        cantidad_excel = None
        if isinstance(fila, dict):
            cantidad_excel = fila.get("cantidad", None)

        _insertar_detalles_b_y_e(page, tabla_index, seed_int, partes, cantidad_excel)


        # ---------------------------------------------
        # BLOQUE DE INFO (debajo de fecha)
        # ---------------------------------------------
        # Izquierda: título del llamado (mismo para todos los ítems)
        if titulo_llamado:
            rect_llamado = fitz.Rect(
                X_LLAMADO,
                y + Y_LLAMADO_OFFSET,
                X_LLAMADO + ANCHO_LLAMADO,
                y + Y_LLAMADO_OFFSET + ALTO_LLAMADO
            )
            insertar_info_autoajustada(
                page,
                rect_llamado,
                str(titulo_llamado).strip(),
                max_lineas=2
            )

        # Derecha: Unidad de medida + Presentación (por ítem)
        texto_info = (
            f"Unidad de medida: {str(unidad_medida).strip()}\n"
            f"Presentación: {str(presentacion).strip()}"
        )

        rect_info = fitz.Rect(
            X_UNI_PRES,
            y + Y_UNI_PRES_OFFSET,
            X_UNI_PRES + ANCHO_UNI_PRES,
            y + Y_UNI_PRES_OFFSET + ALTO_UNI_PRES
        )

        insertar_info_autoajustada(page, rect_info, texto_info, max_lineas=2)

        # -------------------------------
        # DESCRIPCIÓN (NO TOCAR)
        # -------------------------------
        rect = fitz.Rect(
            X_DESCRIPCION,
            y,
            X_DESCRIPCION + ANCHO_DESCRIPCION,
            y + ALTO_BLOQUE
        )

        insertar_texto_autoajustado(page, rect, texto)


        # ---------------------------------------------
        # NUEVO: TEXTOS DE EQUIPOS / MANO DE OBRA / MATERIALES / TRANSPORTE
        # (Solo imprime texto; NO altera cálculos ni coordenadas de números)
        # ---------------------------------------------
        # Estos campos se agregan en app.py usando match_utils.aplicar_match_a_filas().
        textos_partes = [
            fila.get("texto_equipos", ""),
            fila.get("texto_mano_obra", ""),
            fila.get("texto_materiales", ""),
            fila.get("texto_transporte", ""),
        ]

        for idx_parte, t in enumerate(textos_partes):
            t = str(t or "").strip()

            # Equipos y Transporte siempre deben existir: si por algún motivo vienen vacíos, ponemos fallback.
            if idx_parte == 0 and not t:
                t = "Herramientas de mano"
            if idx_parte == 3 and not t:
                t = "Transporte terrestre"

            # Mano de obra y Materiales pueden quedar vacíos según reglas del ítem.
            if not t and idx_parte in (1, 2):
                continue

            x0 = X_CAJAS_PARTES[idx_parte] + PAD_X_PARTES
            y0 = y + Y_CAJAS_PARTES_OFFSET[idx_parte] + PAD_Y_PARTES
            x1 = X_CAJAS_PARTES[idx_parte] + ANCHO_CAJAS_PARTES[idx_parte] - PAD_X_PARTES
            y1 = y + Y_CAJAS_PARTES_OFFSET[idx_parte] + ALTO_CAJAS_PARTES[idx_parte] - PAD_Y_PARTES

            rect_parte = fitz.Rect(x0, y0, x1, y1)
            insertar_texto_partes_autoajustado(page, rect_parte, t)

        # ===============================
        # OPCIÓN B (tapado):
        # Si total es impar y estamos en el ÚLTIMO ítem,
        # y ese ítem está en la PRIMERA tabla de la hoja (posicion_en_hoja == 0),
        # entonces la segunda tabla NO debe verse.
        # ===============================
        if (
            TAPAR_SEGUNDA_TABLA_EN_ULTIMA_HOJA_SI_IMPAR
            and (total % 2 == 1)
            and (i == total - 1)
            and (posicion_en_hoja == 0)
        ):
            # Rectángulo blanco cubriendo desde TAPAR_Y0 hasta el final de la página
            page.draw_rect(
                fitz.Rect(0, TAPAR_Y0, page.rect.width, page.rect.height),
                color=(1, 1, 1),
                fill=(1, 1, 1)
            )

    doc.save(OUTPUT)
    doc.close()
    template_doc.close()

    return OUTPUT

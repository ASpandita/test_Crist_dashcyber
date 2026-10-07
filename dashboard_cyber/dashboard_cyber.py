"""
ETL del Dashboard Cyberday — órdenes B2C ingresadas al WMS vs. target de
venta del cliente, y unidades confirmadas.

Fuentes (se complementan entre sí):
  - BaseWaveDetail.xlsx: todas las órdenes emitidas por el cliente que siguen
    vivas en el WMS (status 0,1,2,3,5). Grano: orden x Sku x Loc. Cuando una
    orden se despacha (status 9) DESAPARECE de este reporte.
  - Fillrate.xlsx: lo confirmado (status 5 y 9) en el rango de fechas que
    pedimos al descargar. Grano: orden x GroupSku. Incluye órdenes en status
    5 que todavía siguen en BaseWaveDetail (se solapan).
  - target_cliente.csv: lo que el cliente dice que va a vender, por día.

Unión sin doble conteo: por (OrderKey, Categoria). Verificado contra datos
reales: en las órdenes que están en ambos reportes, sum(Cantidad) de BWD ==
QtyOriginal de Fillrate y la fecha de ingreso coincide. Si la orden está en
Fillrate, manda Fillrate (trae la confirmación); si no, BaseWaveDetail.

Historial acumulado (historial_ordenes_b2c.csv): cada corrida se fusiona con
lo ya visto, así una orden nunca "desaparece" del dashboard aunque salga del
BWD y quede fuera del rango del Fillrate. Además conserva la hora exacta de
ingreso (AddDate del BWD), que el Fillrate no trae (DataEntryWMS es solo fecha).
"""

import argparse
import base64
import json
import time
import warnings
import zipfile
from datetime import date, datetime

import pandas as pd
from pathlib import Path

# Los Excel que exporta el portal no traen estilo por defecto; openpyxl avisa
# en cada lectura sin que afecte los datos.
warnings.filterwarnings("ignore", message="Workbook contains no default style")

BASE_DIR = Path(__file__).parent
RUTA_BWD = BASE_DIR / "BaseWaveDetail.xlsx"
RUTA_FILLRATE = BASE_DIR / "Fillrate.xlsx"
RUTA_TARGET = BASE_DIR / "target_cliente.csv"
RUTA_HISTORIAL = BASE_DIR / "historial_ordenes_b2c.csv"
RUTA_SALIDA_XLSX = BASE_DIR / "cyber_powerbi.xlsx"
RUTA_PLANTILLA_HTML = BASE_DIR / "dashboard_template.html"
RUTA_SALIDA_HTML = BASE_DIR / "cyber_dashboard.html"
RUTA_SALIDA_ZIP = BASE_DIR / "cyber_dashboard.zip"
RUTA_LOGO_PUMA = BASE_DIR / "logo_puma.png"
RUTA_LOGO_MAERSK = BASE_DIR / "logo_maersk.jpg"

# Solo cuentan las órdenes con ingreso al WMS (DataEntryWMS / AddDate) desde
# este día; deja fuera órdenes previas que siguen abiertas o se confirman
# durante el evento.
FECHA_INICIO_EVENTO = date(2026, 10, 1)
# Rango que se pide al Fill Rate (filtra por fecha de confirmación). Parte un
# día antes del evento por si quedó algo confirmado en el borde del 01/10;
# esas órdenes igual se descartan después por FECHA_INICIO_EVENTO.
FILLRATE_DESDE = date(2026, 9, 30)
FILLRATE_HASTA = date(2026, 10, 16)
INTENTOS_DESCARGA = 3

TIPO_ORDEN = "B2C"
# TargetUnidades: venta/emisión que proyecta el cliente. TargetConfirmaciones:
# unidades que nos comprometemos a confirmar (status 5) ese día.
COLUMNAS_TARGET = ["TargetUnidades", "TargetConfirmaciones"]
CLAVE = ["OrderKey", "Categoria"]

FLAG_LABEL = {"S": "Single", "M": "Multi"}

STATUS_LABEL = {
    "0": "0 · Sin asignar",
    "1": "1 · Asig. parcial",
    "2": "2 · Asignada",
    "3": "3 · En proceso",
    "5": "5 · Confirmada",
    "9": "9 · Despachada",
}


def _normalizar_categoria(s):
    # BWD dice ACCESSORIES y Fillrate ACCESORIES: se unifica para el cruce.
    return s.astype(str).str.strip().str.upper().replace({"ACCESSORIES": "ACCESORIES"})


# ---------------------------------------------------------------------------
# 1. Carga de reportes
# ---------------------------------------------------------------------------
def _columna_flag(df):
    # El nombre puede variar entre la versión manual y la oficial del reporte
    # (Ecom_Single_Flag / ECOM_SingleFlag): se busca ignorando mayúsculas y "_".
    for c in df.columns:
        if str(c).replace("_", "").lower() == "ecomsingleflag":
            return c
    return None


def cargar_bwd(ruta):
    # Fila 0 es el título del reporte; el encabezado real está en la fila 1.
    df = pd.read_excel(ruta, header=1, dtype={"OrderKey": str, "Status": str})
    df = df[df["Type"] == TIPO_ORDEN].copy()
    df["Categoria"] = _normalizar_categoria(df["SkuGroup"])

    col_flag = _columna_flag(df)
    if col_flag is None:
        print("Aviso: BaseWaveDetail sin columna Ecom_Single_Flag — Single/Multi se infiere por unidades.")
        df["FlagSingle"] = pd.NA
    else:
        df["FlagSingle"] = df[col_flag].astype(str).str.strip().str.upper().map(FLAG_LABEL)

    # Status y FlagSingle son a nivel orden (iguales en todas sus líneas):
    # max()/first() solo los toman.
    return df.groupby(CLAVE, as_index=False).agg(
        UnidadesBWD=("Cantidad", "sum"),
        FechaHoraIngreso=("AddDate", "min"),
        StatusBWD=("Status", "max"),
        FlagSingle=("FlagSingle", "first"),
    )


def cargar_fillrate(ruta):
    df = pd.read_excel(ruta, dtype={"Orderkey": str})
    df = df[df["TypeOrd"] == TIPO_ORDEN].copy()
    df = df.rename(columns={"Orderkey": "OrderKey"})
    df["Categoria"] = _normalizar_categoria(df["GroupSku"])
    return df.groupby(CLAVE, as_index=False).agg(
        UnidadesFill=("QtyOriginal", "sum"),
        UnidadesConfirmadasFill=("QtyConfirm", "sum"),
        FechaIngresoFill=("DataEntryWMS", "min"),
        FechaConfirmacionFill=("DateConfirm", "max"),
        FechaDespacho=("DateShipped", "max"),
    )


def cargar_target(ruta):
    if not ruta.exists():
        print(f"Aviso: falta {ruta.name} — el dashboard se genera sin target.")
        return pd.DataFrame(columns=["Fecha"] + COLUMNAS_TARGET)
    df = pd.read_csv(ruta, sep=";")
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True).dt.normalize()
    for c in COLUMNAS_TARGET:
        # Celda vacía = sin target ese día (distinto de 0, que es target cero).
        df[c] = pd.to_numeric(df[c], errors="coerce") if c in df else pd.NA
    return df[["Fecha"] + COLUMNAS_TARGET]


# ---------------------------------------------------------------------------
# 2. Foto actual unificada (BWD + Fillrate)
# ---------------------------------------------------------------------------
def unificar(df_bwd, df_fill, fecha_corrida):
    df = df_bwd.merge(df_fill, on=CLAVE, how="outer")
    en_fill = df["UnidadesFill"].notna()

    df["Unidades"] = df["UnidadesFill"].fillna(df["UnidadesBWD"])
    df["FechaIngreso"] = df["FechaHoraIngreso"].dt.normalize().fillna(df["FechaIngresoFill"])

    # Una orden entra al Fillrate al confirmarse (status 5) y se queda ahí al
    # despacharse (9), pero al pasar a 9 sale del BWD. Por eso: en Fillrate y
    # en BWD = 5; en Fillrate y ya no en BWD = 9. (DateShipped del Fillrate
    # casi siempre viene vacío, no sirve para detectar el despacho.)
    en_bwd = df["UnidadesBWD"].notna()
    df["Status"] = df["StatusBWD"]
    df.loc[en_fill & en_bwd, "Status"] = "5"
    df.loc[en_fill & ~en_bwd, "Status"] = "9"

    df["UnidadesConfirmadas"] = df["UnidadesConfirmadasFill"]
    df["FechaConfirmacion"] = df["FechaConfirmacionFill"]
    df["FuenteConfirmacion"] = pd.Series(pd.NA, index=df.index, dtype="object")
    df.loc[en_fill, "FuenteConfirmacion"] = "Fillrate"

    # Status 5 en BWD cuya ORDEN aún no aparece en Fillrate (confirmada después
    # del corte del reporte): se cuenta como confirmada con fecha = día de
    # corrida hasta que el Fillrate la traiga con su fecha real. Se mira a
    # nivel orden: si la orden ya está en el Fillrate pero le falta esta
    # categoría, no es "pendiente de reporte" sino un quiebre (ver
    # completar_confirmacion_y_quiebre).
    orden_en_fill = df["OrderKey"].isin(df_fill["OrderKey"])
    s5_sin_fill = ~orden_en_fill & (df["StatusBWD"] == "5")
    df.loc[s5_sin_fill, "UnidadesConfirmadas"] = df.loc[s5_sin_fill, "UnidadesBWD"]
    df.loc[s5_sin_fill, "FechaConfirmacion"] = pd.Timestamp(fecha_corrida)
    df.loc[s5_sin_fill, "FuenteConfirmacion"] = "BWD (estimada)"

    df["Quiebre"] = pd.NA  # se calcula después de fusionar con el historial
    df["EnBWD"] = en_bwd
    df["UltimaVezVisto"] = pd.Timestamp(datetime.now()).floor("s")

    # El flag solo viene en el BWD; las órdenes que solo están en el Fillrate
    # lo traen del historial (fusionar_historial) o se infieren después
    # (completar_unidades_y_flag).
    df["FuenteFlag"] = pd.Series(pd.NA, index=df.index, dtype="object")
    df.loc[df["FlagSingle"].notna(), "FuenteFlag"] = "BWD"

    return df[
        CLAVE + [
            "Unidades", "UnidadesBWD", "FechaIngreso", "FechaHoraIngreso", "Status",
            "UnidadesConfirmadas", "Quiebre", "FechaConfirmacion", "FechaDespacho",
            "FuenteConfirmacion", "FlagSingle", "FuenteFlag", "EnBWD", "UltimaVezVisto",
        ]
    ]


# ---------------------------------------------------------------------------
# 3. Historial acumulado
# ---------------------------------------------------------------------------
COLUMNAS_FECHA = ["FechaIngreso", "FechaHoraIngreso", "FechaConfirmacion", "FechaDespacho", "UltimaVezVisto"]


def cargar_historial(ruta):
    if not ruta.exists():
        return None
    df = pd.read_csv(ruta, sep=";", dtype={"OrderKey": str, "Status": str}, parse_dates=COLUMNAS_FECHA)
    if "UnidadesBWD" not in df:
        # Migración única (historial anterior a guardar UnidadesBWD): las
        # órdenes vistas en el BWD (tienen hora de ingreso) guardaron sus
        # unidades del BWD. Válido porque al 02/10 no hay confirmaciones
        # parciales B2C: Fillrate y BWD coinciden en las 737 órdenes cruzadas.
        df["UnidadesBWD"] = df["Unidades"].where(df["FechaHoraIngreso"].notna())
    return df


def fusionar_historial(df_hist, df_actual):
    if df_hist is None or df_hist.empty:
        return df_actual.copy()

    actual = df_actual.set_index(CLAVE)
    hist = df_hist.set_index(CLAVE)

    # Órdenes del historial que ya no están en ningún reporte: si estaban en
    # status 5 se asume despachada (9), que es la única razón normal para salir
    # del BWD. Las que estaban en 0-3 se dejan como estaban (posible anulación;
    # quedan visibles en el historial con EnBWD=False).
    ausentes = hist.index.difference(actual.index)
    hist.loc[ausentes, "EnBWD"] = False
    salio_confirmada = hist.index.isin(ausentes) & (hist["Status"] == "5")
    hist.loc[salio_confirmada, "Status"] = "9"

    # La foto actual manda; el historial rellena lo que falte (filas que ya
    # salieron de los reportes y la hora exacta de ingreso, que solo trae BWD).
    fusion = actual.combine_first(hist)
    # UnidadesBWD = lo pedido originalmente. Se guarda el máximo visto: si el
    # BWD llegara a bajar la cantidad (ej. short pick), no se pierde el original.
    if "UnidadesBWD" in hist:
        fusion["UnidadesBWD"] = pd.concat(
            [actual["UnidadesBWD"], hist["UnidadesBWD"]], axis=1
        ).max(axis=1).reindex(fusion.index)
    return fusion.reset_index()[df_actual.columns]


def completar_unidades_y_flag(df):
    df = df.copy()

    # Emitido = unidades del BWD (lo que pidió el cliente). El Fillrate puede
    # reflejar solo lo confirmado, así que se usa solo si la orden nunca se
    # vio en el BWD.
    df["Unidades"] = df["UnidadesBWD"].combine_first(df["Unidades"])

    # Flag faltante (el reporte aún no trae Ecom_Single_Flag, u orden vista
    # antes de que existiera la columna): regla de negocio por unidades del
    # BWD — 1 unidad = Single, más de 1 = Multi. NUNCA con unidades del
    # Fillrate: una Multi confirmada parcial (2 pedidas, 1 confirmada) se
    # vería como Single. Si la orden nunca pasó por el BWD, queda "Sin flag".
    # Cuando el BWD traiga el flag real, la foto actual lo reemplaza.
    # Solo el flag real del BWD se conserva; los deducidos se recalculan en
    # cada corrida para que siempre reflejen la regla y las unidades vigentes.
    falta = df["FuenteFlag"] != "BWD"
    if falta.any():
        unidades_bwd_orden = df.groupby("OrderKey")["UnidadesBWD"].transform(lambda s: s.sum(min_count=1))
        con_bwd = falta & unidades_bwd_orden.notna()
        df.loc[con_bwd, "FlagSingle"] = (unidades_bwd_orden[con_bwd] <= 1).map({True: "Single", False: "Multi"})
        df.loc[con_bwd, "FuenteFlag"] = "Inferido (unidades BWD)"
        sin_bwd = falta & unidades_bwd_orden.isna()
        df.loc[sin_bwd, "FlagSingle"] = "Sin flag"
        df.loc[sin_bwd, "FuenteFlag"] = "Sin dato (nunca vista en BWD)"
    return df


FUENTE_SIN_CONFIRMAR = "Fillrate (categoría sin confirmar)"


def completar_confirmacion_y_quiebre(df):
    # Quiebre = unidades pedidas (BWD) que no se confirmaron, SOLO en órdenes
    # ya cerradas (status 5 o 9). Dos formas de quiebre:
    #   a) parcial: la categoría viene en el Fillrate con menos unidades
    #      confirmadas que las pedidas (2 pedidas, 1 confirmada = 1 quiebre);
    #   b) total: la orden está en el Fillrate pero esta categoría no aparece
    #      (no se confirmó nada de ella). Se marca con 0 confirmadas, la fecha
    #      de confirmación de la orden y el status de la orden.
    # Se hace después de la fusión porque la fila sin confirmar puede venir
    # solo del historial (la orden ya salió del BWD).
    df = df.copy()
    es_fill = df["FuenteConfirmacion"] == "Fillrate"
    orden_en_fill = es_fill.groupby(df["OrderKey"]).transform("any")
    sin_confirmar = orden_en_fill & ~es_fill
    if sin_confirmar.any():
        fecha_orden = df["FechaConfirmacion"].where(es_fill).groupby(df["OrderKey"]).transform("max")
        status_orden = df["Status"].where(es_fill).groupby(df["OrderKey"]).transform("max")
        df.loc[sin_confirmar, "UnidadesConfirmadas"] = 0
        df.loc[sin_confirmar, "FechaConfirmacion"] = fecha_orden[sin_confirmar]
        df.loc[sin_confirmar, "Status"] = status_orden[sin_confirmar]
        df.loc[sin_confirmar, "FuenteConfirmacion"] = FUENTE_SIN_CONFIRMAR

    cerrada = df["Status"].isin(["5", "9"])
    df["Quiebre"] = (df["Unidades"] - df["UnidadesConfirmadas"].fillna(0)).clip(lower=0).where(cerrada, 0)
    return df


# ---------------------------------------------------------------------------
# 4. Agregados para el dashboard
# ---------------------------------------------------------------------------
def resumir_por_dia(df, df_target):
    emitido = df.groupby("FechaIngreso").agg(
        Emitido=("Unidades", "sum"), OrdenesEmitidas=("OrderKey", "nunique")
    )
    por_tipo = (
        df.groupby(["FechaIngreso", "FlagSingle"])["OrderKey"].nunique()
        .unstack().reindex(columns=["Single", "Multi"]).add_prefix("Ordenes")
    )
    emitido = emitido.join(por_tipo)
    df_conf = df.dropna(subset=["FechaConfirmacion"])
    confirmado = df_conf.groupby("FechaConfirmacion").agg(
        Confirmado=("UnidadesConfirmadas", "sum"), Quiebre=("Quiebre", "sum")
    )
    # Órdenes confirmadas = con al menos 1 unidad confirmada (una orden con
    # quiebre total en todas sus categorías no cuenta como confirmada).
    confirmado = confirmado.join(
        df_conf[df_conf["UnidadesConfirmadas"] > 0].groupby("FechaConfirmacion")["OrderKey"].nunique()
        .rename("OrdenesConfirmadas")
    ).join(df_conf[df_conf["Quiebre"] > 0].groupby("FechaConfirmacion")["OrderKey"].nunique().rename("OrdenesConQuiebre"))
    conf_tipo = (
        df_conf.groupby(["FechaConfirmacion", "FlagSingle"])["OrderKey"].nunique()
        .unstack().reindex(columns=["Single", "Multi"]).add_prefix("OrdenesConf")
    )
    confirmado = confirmado.join(conf_tipo)
    target = df_target.set_index("Fecha")[COLUMNAS_TARGET]
    target = target[target.index >= pd.Timestamp(FECHA_INICIO_EVENTO)]

    dias = emitido.join(confirmado, how="outer").join(target, how="outer")
    dias.index.name = "Fecha"
    dias = dias.sort_index().reset_index()
    for c in [
        "Emitido", "OrdenesEmitidas", "OrdenesSingle", "OrdenesMulti",
        "Confirmado", "OrdenesConfirmadas", "OrdenesConfSingle", "OrdenesConfMulti",
        "Quiebre", "OrdenesConQuiebre",
    ]:
        dias[c] = dias[c].fillna(0).astype(int)
    return dias


def resumir_por_status(df):
    res = df.groupby("Status").agg(Unidades=("Unidades", "sum"), Ordenes=("OrderKey", "nunique")).reset_index()
    res["Etiqueta"] = res["Status"].map(STATUS_LABEL).fillna(res["Status"])
    return res.sort_values("Status")


def resumir_por_categoria(df):
    return df.groupby("Categoria", as_index=False).agg(
        Emitido=("Unidades", "sum"),
        Confirmado=("UnidadesConfirmadas", lambda s: s.sum()),
        Quiebre=("Quiebre", "sum"),
        Ordenes=("OrderKey", "nunique"),
    ).sort_values("Emitido", ascending=False)


def resumir_por_tipo(df):
    # Conteos en ÓRDENES (no unidades): una orden Multi puede tener varias
    # filas (una por categoría), por eso nunique.
    confirmadas = df[df["UnidadesConfirmadas"] > 0]
    pendientes = df[~df["Status"].isin(["5", "9"])]
    con_quiebre = df[df["Quiebre"] > 0]
    res = pd.DataFrame({
        "OrdenesEmitidas": df.groupby("FlagSingle")["OrderKey"].nunique(),
        "UnidadesEmitidas": df.groupby("FlagSingle")["Unidades"].sum(),
        "OrdenesConfirmadas": confirmadas.groupby("FlagSingle")["OrderKey"].nunique(),
        "UnidadesConfirmadas": confirmadas.groupby("FlagSingle")["UnidadesConfirmadas"].sum(),
        "OrdenesPendientes": pendientes.groupby("FlagSingle")["OrderKey"].nunique(),
        "UnidadesQuiebre": df.groupby("FlagSingle")["Quiebre"].sum(),
        "OrdenesConQuiebre": con_quiebre.groupby("FlagSingle")["OrderKey"].nunique(),
        "OrdenesInferidas": df[df["FuenteFlag"].str.startswith("Inferido", na=False)]
        .groupby("FlagSingle")["OrderKey"].nunique(),
    })
    # "Sin flag" solo aparece si hay órdenes que nunca pasaron por el BWD.
    tipos = ["Single", "Multi"] + (["Sin flag"] if "Sin flag" in res.index else [])
    res = res.reindex(tipos).fillna(0).astype(int)
    res.index.name = "Tipo"
    return res.reset_index()


def resumir_quiebres(df):
    # Totales sobre órdenes cerradas (5/9) para el fill rate, y detalle por
    # orden de las que tienen quiebre, para que operación pueda revisarlas.
    cerradas = df[df["Status"].isin(["5", "9"])]
    resumen = {
        "OrdenesCerradas": int(cerradas["OrderKey"].nunique()),
        "UnidadesPedidasCerradas": int(cerradas["Unidades"].sum()),
        "UnidadesConfirmadas": int(cerradas["UnidadesConfirmadas"].fillna(0).sum()),
        "Quiebre": int(cerradas["Quiebre"].sum()),
        "OrdenesConQuiebre": int(cerradas.loc[cerradas["Quiebre"] > 0, "OrderKey"].nunique()),
    }

    ordenes_q = cerradas.loc[cerradas["Quiebre"] > 0, "OrderKey"].unique()
    detalle = cerradas[cerradas["OrderKey"].isin(ordenes_q)].groupby("OrderKey", as_index=False).agg(
        Tipo=("FlagSingle", "first"),
        Status=("Status", "max"),
        FechaIngreso=("FechaIngreso", "min"),
        FechaConfirmacion=("FechaConfirmacion", "max"),
        Pedidas=("Unidades", "sum"),
        Confirmadas=("UnidadesConfirmadas", lambda s: s.fillna(0).sum()),
        Quiebre=("Quiebre", "sum"),
        CategoriasConQuiebre=("Categoria", lambda s: ", ".join(sorted(s[df.loc[s.index, "Quiebre"] > 0]))),
    ).sort_values(["Quiebre", "FechaConfirmacion"], ascending=[False, False])
    for c in ["Pedidas", "Confirmadas", "Quiebre"]:
        detalle[c] = detalle[c].astype(int)
    return resumen, detalle


def resumir_por_hora(df):
    # Solo órdenes con hora exacta (las que el ETL alcanzó a ver en el BWD).
    con_hora = df.dropna(subset=["FechaHoraIngreso"]).copy()
    con_hora["Fecha"] = con_hora["FechaHoraIngreso"].dt.normalize()
    con_hora["Hora"] = con_hora["FechaHoraIngreso"].dt.hour
    claves = ["Fecha", "Hora"]
    total = con_hora.groupby(claves).agg(Unidades=("Unidades", "sum"), Ordenes=("OrderKey", "nunique"))
    por_tipo = con_hora.groupby(claves + ["FlagSingle"]).agg(
        Unidades=("Unidades", "sum"), Ordenes=("OrderKey", "nunique")
    ).unstack().reindex(columns=pd.MultiIndex.from_product([["Unidades", "Ordenes"], ["Single", "Multi"]]))
    por_tipo.columns = [f"{m}{t}" for m, t in por_tipo.columns]  # ej. OrdenesSingle, UnidadesMulti
    return total.join(por_tipo).fillna(0).astype(int).reset_index()


# ---------------------------------------------------------------------------
# 5. Export
# ---------------------------------------------------------------------------
def exportar_excel(df_ordenes, dias, status, categoria, tipo, quiebres_detalle, ruta):
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        dias.to_excel(writer, sheet_name="Resumen_Dia", index=False)
        status.to_excel(writer, sheet_name="Resumen_Status", index=False)
        categoria.to_excel(writer, sheet_name="Resumen_Categoria", index=False)
        tipo.to_excel(writer, sheet_name="Resumen_SingleMulti", index=False)
        quiebres_detalle.to_excel(writer, sheet_name="Quiebres", index=False)
        df_ordenes.to_excel(writer, sheet_name="Ordenes_B2C", index=False)
    print(f"Exportado: {ruta}")


def _a_json(df, columnas_fecha=(), formato="%Y-%m-%d"):
    df = df.copy()
    for c in columnas_fecha:
        df[c] = df[c].dt.strftime(formato)
    df = df.astype(object).where(df.notna(), None)
    return df.to_json(orient="records", force_ascii=False)


def _logo_data_uri(ruta):
    # Incrustado en base64: el HTML viaja solo (zip por correo) sin perder
    # las imágenes. Si falta el archivo, la plantilla oculta el <img>.
    if not ruta.exists():
        return ""
    mime = "image/png" if ruta.suffix.lower() == ".png" else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(ruta.read_bytes()).decode()}"


def generar_dashboard_html(dias, status, categoria, tipo, horas, quiebres, ruta_plantilla, ruta_salida):
    quiebre_resumen, quiebre_detalle = quiebres
    plantilla = ruta_plantilla.read_text(encoding="utf-8")
    html = (
        plantilla
        .replace("__DIAS_JSON__", _a_json(dias, ["Fecha"]))
        .replace("__QUIEBRE_RESUMEN_JSON__", json.dumps(quiebre_resumen))
        .replace("__QUIEBRE_DETALLE_JSON__", _a_json(quiebre_detalle, ["FechaIngreso", "FechaConfirmacion"]))
        .replace("__TIPO_JSON__", _a_json(tipo))
        .replace("__STATUS_JSON__", _a_json(status))
        .replace("__CATEGORIA_JSON__", _a_json(categoria))
        .replace("__HORAS_JSON__", _a_json(horas, ["Fecha"]))
        .replace("__FECHA_GENERACION__", datetime.now().strftime("%Y-%m-%d %H:%M"))
        .replace("__FECHA_HOY__", date.today().isoformat())
        .replace("__LOGO_PUMA__", _logo_data_uri(RUTA_LOGO_PUMA))
        .replace("__LOGO_MAERSK__", _logo_data_uri(RUTA_LOGO_MAERSK))
    )
    ruta_salida.write_text(html, encoding="utf-8")
    print(f"Exportado: {ruta_salida}")


def comprimir_dashboard(ruta_html, ruta_zip):
    # .zip en vez de .html directo: los filtros de correo corporativo suelen
    # bloquear adjuntos .html.
    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(ruta_html, arcname=ruta_html.name)
    print(f"Exportado: {ruta_zip}")


# ---------------------------------------------------------------------------
# 6. Orquestación
# ---------------------------------------------------------------------------
def ejecutar(descargar=True):
    if descargar:
        from descargar_reportes_cyber import descargar_reportes_cyber

        # El portal falla de forma intermitente (timeouts en el diálogo de
        # exportación); un reintento suele bastar.
        for intento in range(1, INTENTOS_DESCARGA + 1):
            try:
                descargar_reportes_cyber(FILLRATE_DESDE, FILLRATE_HASTA, headless=True)
                break
            except Exception as e:
                if intento < INTENTOS_DESCARGA:
                    print(f"Intento {intento} de descarga falló ({type(e).__name__}) — reintentando...")
                else:
                    # En pleno evento es mejor un dashboard con la última data
                    # descargada que no tener dashboard.
                    print(f"Aviso: falló la descarga ({e!r}) — se usan los últimos archivos disponibles.")

    df_actual = unificar(cargar_bwd(RUTA_BWD), cargar_fillrate(RUTA_FILLRATE), date.today())
    df_ordenes = fusionar_historial(cargar_historial(RUTA_HISTORIAL), df_actual)
    df_ordenes = completar_unidades_y_flag(df_ordenes)
    df_ordenes = completar_confirmacion_y_quiebre(df_ordenes)
    df_ordenes.to_csv(RUTA_HISTORIAL, sep=";", index=False)

    del_evento = df_ordenes[df_ordenes["FechaIngreso"] >= pd.Timestamp(FECHA_INICIO_EVENTO)]

    dias = resumir_por_dia(del_evento, cargar_target(RUTA_TARGET))
    status = resumir_por_status(del_evento)
    categoria = resumir_por_categoria(del_evento)
    tipo = resumir_por_tipo(del_evento)
    horas = resumir_por_hora(del_evento)
    quiebres = resumir_quiebres(del_evento)

    exportar_excel(del_evento, dias, status, categoria, tipo, quiebres[1], RUTA_SALIDA_XLSX)
    generar_dashboard_html(dias, status, categoria, tipo, horas, quiebres, RUTA_PLANTILLA_HTML, RUTA_SALIDA_HTML)
    comprimir_dashboard(RUTA_SALIDA_HTML, RUTA_SALIDA_ZIP)


def main():
    parser = argparse.ArgumentParser(description="Dashboard Cyberday B2C")
    parser.add_argument("--sin-descarga", action="store_true", help="usar los Excel ya descargados")
    parser.add_argument("--cada", type=int, default=0, metavar="MIN", help="repetir cada MIN minutos")
    args = parser.parse_args()

    while True:
        ejecutar(descargar=not args.sin_descarga)
        if args.cada <= 0:
            break
        print(f"Próxima actualización en {args.cada} min ({datetime.now():%H:%M})")
        time.sleep(args.cada * 60)


if __name__ == "__main__":
    main()

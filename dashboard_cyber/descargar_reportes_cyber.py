"""
Descarga automática de los 2 reportes que alimentan dashboard_cyber.py,
desde el portal WMS (JInfonet). Mismo flujo base que
dashboard_ocupacion/descargar_inventory_snapshot.py:

  Login -> Public Folder -> carpeta CHL -> buscar reporte -> abrir
  -> (Fillrate: llenar parámetros de fecha) -> Export -> Excel Workbook.

  - BaseWaveDetail: órdenes vivas (status 0,1,2,3,5). Sin parámetros.
  - Fillrate: lo confirmado (status 5 y 9) en un rango de fechas que
    elegimos nosotros (FILLRATE_DESDE / FILLRATE_HASTA en dashboard_cyber.py).

Ojo: el portal trae JS antiguo que rompe algunos motores de selectores de
Playwright (ej. locator("text=/regex/")) y page.evaluate que devuelve
objetos; usar locators simples por id/name/role.
"""

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
# Mismas credenciales que el dashboard de ocupación (no duplicar el secreto).
RUTA_CREDENCIALES = BASE_DIR.parent / "dashboard_ocupacion" / "credenciales_wms.json"

RUTA_BWD = BASE_DIR / "BaseWaveDetail.xlsx"
RUTA_FILLRATE = BASE_DIR / "Fillrate.xlsx"

# Nombres verificados en el portal (CHL/WMS/PUMACL/INVENTORY y .../Operations).
BUSQUEDA_BWD = "PUMACL_baseWavesDetail"
NOMBRE_BWD = "PUMACL_baseWavesDetail.cls"
BUSQUEDA_FILLRATE = "PUMACL_FILL_RATE"
NOMBRE_FILLRATE = "PUMACL_FILL_RATE.cls"

# El formato que acepta el campo de fecha depende del locale de la cuenta/navegador
# (con un usuario aceptaba "October 1, 2026", con otro solo "dd/MM/y"). Se usa
# dd/MM/yyyy, que es numérico y no depende del idioma de los meses.
def _fecha_portal(d):
    return d.strftime("%d/%m/%Y")


def cargar_credenciales(ruta):
    return json.loads(ruta.read_text(encoding="utf-8"))


def _login(page, cred):
    page.goto(cred["url_login"])
    page.wait_for_timeout(1500)
    page.locator('[id="jrs.auth_uid"]').click()
    page.locator('[id="jrs.auth_uid"]').fill(cred["usuario"])
    page.locator('[id="jrs.auth_uid"]').press("Tab")
    page.locator('input[name="jrs.auth_pwd"]').fill(cred["password"])
    page.locator('input[name="jrs.auth_pwd"]').press("Enter")
    page.wait_for_timeout(3000)


def _ir_a_carpeta_chl(page):
    # Siempre se navega desde Public Folder: un navegador nuevo de Playwright
    # no tiene la sesión "recordada" que reabre el último reporte. Se hace una
    # sola vez: después del primer reporte "Public Folder" ya no es visible,
    # pero el buscador de CHL sigue disponible para el siguiente.
    page.get_by_text("Public Folder").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)
    page.frame(name="main").get_by_role("link", name="CHL", exact=True).click()
    page.wait_for_timeout(3000)


def _abrir_reporte(page, context, busqueda, nombre_cls):
    listado = page.frame(name="main")
    buscador = listado.get_by_placeholder("Search")
    buscador.click()
    buscador.fill(busqueda)
    buscador.press("Enter")
    page.wait_for_timeout(4000)

    listado = page.frame(name="main")
    with context.expect_page(timeout=15000) as popup_info:
        listado.get_by_text(nombre_cls, exact=True).click()
    reporte = popup_info.value
    reporte.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    return reporte


def llenar_parametros_fillrate(reporte, fecha_desde, fecha_hasta):
    # Pantalla "Enter Parameter Values" (en el documento del popup, no en un
    # iframe). Se escribe en el campo visible y Tab para que el portal lo
    # valide y copie al campo oculto que realmente se envía.
    reporte.wait_for_timeout(3000)
    for param, valor in (("ParamFechaDesde", fecha_desde), ("ParamFechaHasta", fecha_hasta)):
        campo = reporte.locator(f'[id="input_jrs.param${param}"]')
        campo.click()
        campo.fill(_fecha_portal(valor))
        campo.press("Tab")
        reporte.wait_for_timeout(500)
    reporte.locator('input[name="Submit_Btn2"]').click()
    reporte.wait_for_load_state("networkidle")
    reporte.wait_for_timeout(5000)


def _exportar_excel(reporte, nombre_archivo, ruta_destino):
    # Esperas explícitas: el portal es lento entre pasos y el reporte puede
    # tardar minutos en generarse (mismo criterio que el de ocupación).
    reporte.get_by_role("img", name="Export").click(timeout=300000)
    reporte.wait_for_timeout(2000)

    frame = reporte.locator("#client iframe").content_frame
    frame.get_by_role("cell", name="File Name:").click()
    frame.locator("#labFileName").fill(nombre_archivo)
    frame.get_by_role("button", name="Submit").click()
    reporte.wait_for_timeout(3000)

    frame.get_by_role("link", name="excelExcel").click(timeout=60000)
    reporte.wait_for_timeout(1500)
    frame.locator("#i_ver").select_option("1")
    reporte.wait_for_timeout(1500)

    with reporte.expect_download(timeout=240000) as download_info:
        frame.get_by_role("button", name="OK").click()
    download_info.value.save_as(ruta_destino)
    print(f"Descargado: {ruta_destino}")


def descargar_reportes_cyber(fillrate_desde, fillrate_hasta, headless=True):
    cred = cargar_credenciales(RUTA_CREDENCIALES)

    with sync_playwright() as playwright:
        # Se usa el Edge instalado: el proxy corporativo bloquea la descarga
        # del Chromium de Playwright (UNABLE_TO_GET_ISSUER_CERT_LOCALLY).
        browser = playwright.chromium.launch(headless=headless, channel="msedge")
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        _login(page, cred)
        _ir_a_carpeta_chl(page)

        reporte = _abrir_reporte(page, context, BUSQUEDA_BWD, NOMBRE_BWD)
        _exportar_excel(reporte, "BaseWaveDetail", RUTA_BWD)
        reporte.close()

        reporte = _abrir_reporte(page, context, BUSQUEDA_FILLRATE, NOMBRE_FILLRATE)
        llenar_parametros_fillrate(reporte, fillrate_desde, fillrate_hasta)
        _exportar_excel(reporte, "Fillrate", RUTA_FILLRATE)
        reporte.close()

        context.close()
        browser.close()


if __name__ == "__main__":
    from dashboard_cyber import FILLRATE_DESDE, FILLRATE_HASTA

    descargar_reportes_cyber(FILLRATE_DESDE, FILLRATE_HASTA, headless=False)

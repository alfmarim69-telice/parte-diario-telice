"""
cargar_costes_sp.py
───────────────────
Lee el Excel Maestro de Obra y actualiza la lista M_Partidas en SharePoint
añadiendo los campos CosteUd (coste unitario) y CosteTot (coste total).

PREREQUISITOS:
  pip install openpyxl msal requests --break-system-packages

CONFIGURACIÓN (rellenar antes de ejecutar):
  - CLIENT_ID, TENANT_ID: del App Registration de Azure (los mismos que usa la app)
  - SITE_URL: URL del sitio SharePoint donde está M_Partidas
  - EXCEL_PATH: ruta al archivo Excel Maestro de Obra

CAMPOS QUE ESCRIBE EN M_Partidas:
  - CosteUd  → Coste UD Proyecto (col K del Excel)
  - CosteTot → TOTAL Coste       (col AC del Excel)

Si esos campos no existen en la lista, el script avisa y para.
Créalos primero en SharePoint (tipo Número, decimales permitidos).
"""

import json
import sys
import time
from openpyxl import load_workbook

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
CLIENT_ID  = "TU_CLIENT_ID"      # App Registration → Overview → Application (client) ID
TENANT_ID  = "TU_TENANT_ID"      # Azure AD → Overview → Tenant ID / Directory ID
SITE_URL   = "https://telice.sharepoint.com/sites/ParteDiario"   # ajustar si difiere
EXCEL_PATH = "23-039_Maestro_Obra_13_DEF.xlsx"  # ruta relativa o absoluta

# Nombre interno de la lista en SharePoint (exactamente como está)
LISTA = "M_Partidas"

# Hoja del Excel y columnas (0-indexed)
HOJA         = "Partidas_Produccion"
COL_CODIGO   = 0   # A  → Código partida (Title en SP)
COL_MEDICION = 8   # I  → Medición (para filtrar filas vacías)
COL_COSTE_UD = 10  # K  → Coste UD Proyecto
COL_COSTE_TOT= 28  # AC → TOTAL Coste
FILA_DATOS   = 5   # primera fila con datos reales (1-indexed, filas 1-4 son título/cabecera/caps)

# ─── AUTENTICACIÓN (Device Code Flow — no necesita secreto) ──────────────────
try:
    import msal
except ImportError:
    print("ERROR: instala msal → pip install msal --break-system-packages")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("ERROR: instala requests → pip install requests --break-system-packages")
    sys.exit(1)


def get_token():
    """Obtiene token de acceso usando Device Code Flow (abre el navegador)."""
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    scopes = ["https://graph.microsoft.com/Sites.ReadWrite.All"]

    # Intentar token en caché primero
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print("✓ Token obtenido desde caché")
            return result["access_token"]

    # Device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise ValueError(f"Error iniciando device flow: {flow}")

    print("\n" + "="*60)
    print("AUTENTICACIÓN REQUERIDA")
    print(f"  1. Abre: {flow['verification_uri']}")
    print(f"  2. Introduce el código: {flow['user_code']}")
    print("="*60 + "\n")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise ValueError(f"Error obteniendo token: {result.get('error_description', result)}")

    print("✓ Autenticación correcta\n")
    return result["access_token"]


# ─── FUNCIONES GRAPH API ──────────────────────────────────────────────────────

def get_site_id(token, site_url):
    """Obtiene el siteId de Graph a partir de la URL del sitio."""
    # Extraer hostname y path del site_url
    # ej. https://telice.sharepoint.com/sites/ParteDiario
    parts = site_url.replace("https://", "").split("/", 1)
    hostname = parts[0]
    path = parts[1] if len(parts) > 1 else ""
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    site_id = r.json()["id"]
    print(f"✓ Site ID: {site_id}")
    return site_id


def get_lista_id(token, site_id, nombre_lista):
    """Obtiene el ID interno de la lista en SP."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$filter=displayName eq '{nombre_lista}'"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    items = r.json().get("value", [])
    if not items:
        raise ValueError(f"Lista '{nombre_lista}' no encontrada en el sitio.")
    lista_id = items[0]["id"]
    print(f"✓ Lista '{nombre_lista}' ID: {lista_id}")
    return lista_id


def get_all_items(token, site_id, lista_id):
    """Lee todos los items de la lista (paginado)."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items?$select=id,fields/Title&$top=500"
    items = []
    while url:
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    print(f"✓ {len(items)} items leídos de '{LISTA}'")
    return items


def patch_item(token, site_id, lista_id, item_id, campos):
    """Actualiza campos de un item existente."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista_id}/items/{item_id}"
    r = requests.patch(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps({"fields": campos})
    )
    if r.status_code not in (200, 204):
        raise ValueError(f"PATCH {item_id} → {r.status_code}: {r.text[:200]}")


# ─── LEER EXCEL ───────────────────────────────────────────────────────────────

def leer_costes_excel(path):
    """Lee los costes unitarios y totales del Excel Maestro."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[HOJA]
    costes = {}
    filas_vacias = 0

    for row in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        cod      = row[COL_CODIGO]
        medicion = row[COL_MEDICION]
        coste_ud = row[COL_COSTE_UD]
        coste_tot= row[COL_COSTE_TOT]

        # Saltar filas de capítulo/título (sin medición) y filas sin código
        if not cod or not isinstance(cod, str) or not medicion:
            filas_vacias += 1
            continue

        costes[cod.strip()] = {
            "CosteUd":  round(float(coste_ud),  6) if coste_ud  is not None else 0.0,
            "CosteTot": round(float(coste_tot), 2) if coste_tot is not None else 0.0,
        }

    wb.close()
    print(f"✓ {len(costes)} partidas leídas del Excel (saltadas {filas_vacias} filas sin datos)")
    return costes


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=== Carga de CosteUd y CosteTot en M_Partidas (SharePoint) ===\n")

    # 1. Leer Excel
    print(f"Leyendo {EXCEL_PATH}...")
    costes = leer_costes_excel(EXCEL_PATH)
    if not costes:
        print("ERROR: No se encontraron partidas en el Excel.")
        sys.exit(1)

    # Preview
    preview = list(costes.items())[:3]
    print("Preview:")
    for cod, vals in preview:
        print(f"  {cod}: CosteUd={vals['CosteUd']}, CosteTot={vals['CosteTot']}")
    print()

    # 2. Obtener token
    token = get_token()

    # 3. Obtener IDs de sitio y lista
    site_id  = get_site_id(token, SITE_URL)
    lista_id = get_lista_id(token, site_id, LISTA)

    # 4. Leer items existentes en SP → mapa Title → item_id
    items_sp = get_all_items(token, site_id, lista_id)
    mapa_sp  = {item["fields"]["Title"]: item["id"] for item in items_sp}

    # 5. Actualizar cada partida
    actualizados = 0
    no_encontrados = []
    errores = []

    total = len(costes)
    print(f"\nActualizando {total} partidas en SharePoint...")

    for i, (cod, vals) in enumerate(costes.items(), 1):
        if cod not in mapa_sp:
            no_encontrados.append(cod)
            continue

        try:
            patch_item(token, site_id, lista_id, mapa_sp[cod], {
                "CosteUd":  vals["CosteUd"],
                "CosteTot": vals["CosteTot"],
            })
            actualizados += 1
            if i % 20 == 0 or i == total:
                print(f"  {i}/{total} ({actualizados} actualizados, {len(errores)} errores)...")
            # Pequeña pausa para no saturar Graph API (throttling)
            time.sleep(0.05)

        except Exception as e:
            errores.append(f"{cod}: {e}")

    # 6. Resumen
    print("\n=== RESUMEN ===")
    print(f"  ✓ Actualizados:    {actualizados}")
    print(f"  ⚠ No encontrados:  {len(no_encontrados)}")
    print(f"  ✗ Errores:         {len(errores)}")

    if no_encontrados:
        print(f"\nPartidas del Excel no encontradas en SP ({len(no_encontrados)}):")
        for cod in no_encontrados[:10]:
            print(f"  - {cod}")
        if len(no_encontrados) > 10:
            print(f"  ... y {len(no_encontrados)-10} más")

    if errores:
        print(f"\nErrores ({len(errores)}):")
        for e in errores[:5]:
            print(f"  - {e}")

    print("\nDone.")


if __name__ == "__main__":
    main()

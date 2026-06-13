import openpyxl
import pandas as pd
import requests
import datetime
import os
import subprocess
import re

# Configuración del Google Sheet
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/17r2QvAAQsZZK7YB_UXstmDObZC7WkofGII_cbLUVybU/export?format=xlsx"
LOCAL_XLSX_PATH = "google_sheet_temp.xlsx"
DATA_JS_PATH = "data.js"

# Mapeo de días de la semana
DIA_MAP = {
    "lunes": "lun",
    "martes": "mar",
    "miercoles": "mie",
    "miércoles": "mie",
    "jueves": "jue",
    "viernes": "vie",
    "sabado": "sab",
    "sábado": "sab",
    "domingo": "dom"
}

# Mapeo de meses según pestañas
MES_MAP = {
    "Mayo25": "may",
    "JUN25": "jun",
    "JUL25": "jul",
    "Agos25": "ago",
    "SEP25": "sep",
    "Oct25": "oct",
    "Nov25": "nov",
    "Dic25": "dic",
    "Ene26": "ene",
    "Feb26": "feb",
    "Marz26": "mar",
    "Abril26": "abr",
    "May26": "may",
    "Jun26": "jun"
}

# El orden en el que se listan los meses en el dashboard comparativo
MONTH_ORDER = [
    "Mayo25", "JUN25", "JUL25", "Agos25", "SEP25", "Oct25", "Nov25", "Dic25",
    "Ene26", "Feb26", "Marz26", "Abril26", "May26", "Jun26"
]

# Datos históricos de Outlet extraídos de index.html original
OUTLET_HISTORIC_DATA = [
    # Datos de Febrero 2026
    ["2 feb", None, 39200, -39200],
    ["3 feb", 12000, 35000, -23000],
    ["4 feb", 50580, 42900, 7680],
    ["5 feb", 40000, 40220.68, -220.68],
    ["6 feb", 76900, 35150, 41750],
    ["7 feb", 183000, 40000, 143000],
    ["9 feb", 504698, 35400, 469298],
    ["10 feb", 26750, 105000, -78250],
    ["11 feb", 67500, 35000, 32500],
    ["12 feb", 15000, 1500, 13500],
    ["13 feb", 69000, 44000, 25000],
    ["14 feb", 79200, 0, 79200],
    ["17 feb", 15000, 25000, -10000],
    ["18 feb", 68000, 32000, 36000],
    ["19 feb", 0, 65000, -50000],
    ["20 feb", 15000, 70000, -55000],
    ["21 feb", 97900, 65000, 32900],
    ["23 feb", 40000, 35000, 5000],
    ["24 feb", 91500, 32600, 58900],
    ["25 feb", 184000, 1500, 182500],
    ["26 feb", 186000, 49500, 136500],
    ["27 feb", 171300, 35000, 136300],
    ["28 feb", 118000, 0, 118000],
    # Datos de Marzo 2026
    ["2 mar", 30000, 55000, -25000],
    ["3 mar", 50910, 70400, -19490],
    ["4 mar", 0, 55000, -55000],
    ["5 mar", 30000, 64500, -34500],
    ["6 mar", 0, 60000, -60000],
    ["7 mar", 183000, 40000, 143000],
    ["9 mar", 247500, 60400, 187100],
    ["10 mar", 38000, 96300, -58300],
    ["11 mar", 160000, 63000, 97000],
    ["12 mar", 28000, 58800, -30800],
    ["13 mar", 41100, 66500, -25400],
    ["14 mar", 183000, 0, 183000],
    ["16 mar", 10000, 55000, -45000],
    ["17 mar", 64000, 67500, -3500],
    ["18 mar", 49200, 55000, -5800],
    ["19 mar", 0, 55000, -55000],
    ["21 mar", 0, 40000, -40000],
    ["23 mar", 36000, 65000, -29000],
    ["24 mar", 0, 60000, -60000],
    ["25 mar", 0, 55000, -55000],
    ["26 mar", 64300, 35000, 29300],
    ["27 mar", 44100, 52990, -8890],
    ["28 mar", 0, 40800, -40800],
    ["30 mar", 43000, 55000, -12000],
    ["31 mar", 156200, 56200, 100000]
]

def clean_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "nan" or str(val).lower() == "null":
        return 0
    try:
        s = str(val).strip().replace("$", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except Exception:
        return 0

def is_valid_day_num(val):
    if val is None:
        return False
    try:
        f_val = float(str(val).strip())
        return f_val.is_integer() and 1 <= f_val <= 31
    except ValueError:
        return False

def download_sheet():
    print("Descargando Google Sheet en formato XLSX...")
    r = requests.get(SPREADSHEET_URL)
    if r.status_code == 200:
        with open(LOCAL_XLSX_PATH, "wb") as f:
            f.write(r.content)
        print("Descarga completada.")
        return True
    else:
        print(f"Error al descargar: {r.status_code}")
        return False

def parse_month_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        print(f"Advertencia: Pestaña '{sheet_name}' no encontrada en el Excel.")
        return []
    
    sheet = wb[sheet_name]
    rows = list(sheet.iter_rows(values_only=True))
    
    daily_records = []
    
    current_day_num = None
    current_day_name = ""
    current_turno = "" # "MAÑANA" o "TARDE"
    
    day_data = None
    
    for idx, row in enumerate(rows):
        if not row:
            continue
            
        col_a = str(row[0]).strip().lower() if row[0] is not None else ""
        col_b = row[1]
        col_c = str(row[2]).strip().lower() if row[2] is not None else ""
        col_d = row[3]
        
        # 1. Detectar cambio de día
        if col_a in DIA_MAP and is_valid_day_num(col_b):
            if day_data and day_data["day_num"] is not None:
                daily_records.append(day_data)
                
            current_day_num = int(float(str(col_b).strip()))
            current_day_name = DIA_MAP[col_a]
            current_turno = ""
            
            day_data = {
                "day_num": current_day_num,
                "day_name": current_day_name,
                "month_code": MES_MAP[sheet_name],
                "ventas_tm": 0,
                "gastos_tm": 0,
                "empleado_tm": None,
                "ventas_tt": 0,
                "gastos_tt": 0,
                "empleado_tt": None
            }
            continue
            
        if day_data is None:
            continue
            
        # 2. Detectar Turno
        if "turno maña" in col_a or "primera maña" in col_a:
            current_turno = "MAÑANA"
            continue
        elif "turno tarde" in col_a:
            current_turno = "TARDE"
            continue
            
        # 3. Detectar Cierres / Totales de Ventas y Gastos
        if current_turno == "MAÑANA":
            if any(x in col_c or x in col_a for x in ["venta total tm", "facturacion total tm", "facturación total tm"]):
                day_data["ventas_tm"] = clean_num(col_d)
            elif any(x in col_c or x in col_a for x in ["gastos total tm", "gasto total tm"]):
                day_data["gastos_tm"] = clean_num(col_d)
            elif "sobrente de caja" in col_c or "sobrantede caja" in col_c or "sobrante/faltante" in col_c:
                emp = str(row[0]).strip() if row[0] is not None else ""
                if emp and emp.lower() not in ["", "nan", "null"]:
                    day_data["empleado_tm"] = emp.upper()
                    
        elif current_turno == "TARDE":
            if any(x in col_c or x in col_a for x in ["venta total tt", "facturacion total tt", "facturación total tt", "venta total tm", "facturacion total tm"]):
                day_data["ventas_tt"] = clean_num(col_d)
            elif any(x in col_c or x in col_a for x in ["gastos total tm", "gastos total tt", "gasto total tm", "gastos total del dia", "gastos total del día", "gasto total del dia"]):
                day_data["gastos_tt"] = clean_num(col_d)
            elif "sobrante/faltante" in col_c or "sobrantede caja" in col_c:
                emp = str(row[0]).strip() if row[0] is not None else ""
                if emp and emp.lower() not in ["", "nan", "null"]:
                    day_data["empleado_tt"] = emp.upper()

    if day_data and day_data["day_num"] is not None:
        daily_records.append(day_data)
        
    formatted_rows = []
    for r in daily_records:
        label = f"{r['day_name']} {r['day_num']} {r['month_code']}"
        ventas = r["ventas_tm"] + r["ventas_tt"]
        gastos = r["gastos_tm"] + r["gastos_tt"]
        ganancias = ventas - gastos
        
        # Filtro: si ventas y gastos son cero, saltamos el registro (salvo para el mes en curso)
        if ventas == 0 and gastos == 0 and sheet_name != "Jun26":
            continue
            
        formatted_rows.append([
            label,
            ventas,
            gastos,
            ganancias,
            r["empleado_tm"] or "—",
            r["empleado_tt"] or "—"
        ])
        
    return formatted_rows

def main():
    if not download_sheet():
        print("No se pudo obtener el archivo de Google Sheets. Cancelando sincronización.")
        return
        
    wb = openpyxl.load_workbook(LOCAL_XLSX_PATH, data_only=True)
    
    showroom_data = []
    
    print("Procesando pestañas mensuales...")
    for month_sheet in MONTH_ORDER:
        print(f"Procesando {month_sheet}...")
        sheet_rows = parse_month_sheet(wb, month_sheet)
        showroom_data.extend(sheet_rows)
        print(f"Cargados {len(sheet_rows)} días.")
        
    print(f"Total de registros de Showroom procesados: {len(showroom_data)}")
    
    # Generar el archivo data.js
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    js_content = f"""// Archivo generado automáticamente por el script de sincronización.
// Última actualización: {timestamp}

const DATOS_SINCRONIZADOS = {{
  ultima_actualizacion: "{timestamp}",
  showroom: {str(showroom_data)},
  outlet: {str(OUTLET_HISTORIC_DATA)}
}};
"""
    
    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)
        
    print(f"Archivo '{DATA_JS_PATH}' escrito con éxito.")
    
    if os.path.exists(LOCAL_XLSX_PATH):
        os.remove(LOCAL_XLSX_PATH)
        
    try:
        print("Comprobando cambios en Git...")
        status = subprocess.check_output(["git", "status", "--porcelain"], text=True)
        if "data.js" in status or "package.json" in status or "sync_sheets.py" in status:
            print("Subiendo cambios actualizados a GitHub...")
            subprocess.run(["git", "add", "data.js", "package.json", "sync_sheets.py"])
            subprocess.run(["git", "commit", "-m", f"Sincronización automática de Google Sheets - {timestamp}"])
            subprocess.run(["git", "push", "origin", "main"])
            print("Push a GitHub completado con éxito.")
        else:
            print("No hay cambios en los datos. No se requiere hacer push.")
    except Exception as e:
        print(f"Error al subir cambios a GitHub: {str(e)}")

if __name__ == "__main__":
    main()

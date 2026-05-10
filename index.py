"""
Script: generar_oratoria.py
Dependencias: pip install openpyxl
Uso: python generar_oratoria.py
"""

import os
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

COLOR_HEADER_BG  = "1F3864"
COLOR_HEADER_FG  = "FFFFFF"
COLOR_MENU_TITLE = "1F3864"
COLOR_MENU_FG    = "FFFFFF"
COLOR_LINK_BG    = "2E75B6"
COLOR_LINK_FG    = "FFFFFF"
COLOR_ROW_ALT    = "DCE6F1"
COLOR_BACK_BG    = "C00000"
TALLER_OBJETIVO  = "TALLER DE ORATORIA"

def borde(color="1F3864"):
    lado = Side(style="thin", color=color)
    return Border(left=lado, right=lado, top=lado, bottom=lado)

def limpiar_hoja(nombre):
    # 1. Eliminar tildes
    nombre = unicodedata.normalize("NFD", nombre)
    nombre = "".join(c for c in nombre if unicodedata.category(c) != "Mn")
    # 2. Reemplazar espacios por guiones bajos
    nombre = nombre.replace(" ", "_")
    # 3. Eliminar caracteres inválidos para Excel (incluida la coma)
    for c in ['\\', '/', '*', '?', ':', '[', ']', "'", ',']:
        nombre = nombre.replace(c, '')
    # 4. Mayúsculas y límite de 31 caracteres
    return nombre.upper()[:31]

def autoajustar(ws):
    for col in ws.columns:
        ancho = 0
        for celda in col:
            if celda.value:
                ancho = max(ancho, len(str(celda.value)))
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ancho + 4, 60)

def leer_datos(ruta):
    wb = openpyxl.load_workbook(ruta)
    ws = wb.active
    enc = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    filas = []
    for r in range(2, ws.max_row + 1):
        fila = {enc[c]: ws.cell(r, c + 1).value for c in range(len(enc))}
        filas.append(fila)
    wb.close()
    return enc, filas

def filtrar_y_agrupar(filas, col_taller, col_escuela):
    grupos = {}
    for fila in filas:
        taller = str(fila.get(col_taller, "") or "").strip().upper()
        if taller == TALLER_OBJETIVO.upper():
            escuela = str(fila.get(col_escuela, "SIN_ESCUELA") or "SIN_ESCUELA").strip()
            grupos.setdefault(escuela, []).append(fila)
    return grupos

def crear_hoja_escuela(wb, nombre_escuela, filas, enc, nombre_menu):
    nombre_hoja = limpiar_hoja(nombre_escuela)
    ws = wb.create_sheet(title=nombre_hoja)

    # Botón Volver
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "◀  VOLVER AL MENU"
    c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=COLOR_BACK_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.hyperlink = f"#{nombre_menu}!A1"
    c.border = borde(COLOR_BACK_BG)

    # Título (muestra nombre original con tildes, solo visual)
    ws.row_dimensions[2].height = 10
    ws.row_dimensions[3].height = 36
    n = len(enc)
    ul = get_column_letter(n)
    ws.merge_cells(f"A3:{ul}3")
    t = ws["A3"]
    t.value = f"TALLER DE ORATORIA  —  {nombre_escuela.upper()}"
    t.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=COLOR_MENU_TITLE)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 8

    # Encabezados tabla
    ws.row_dimensions[5].height = 28
    for ci, col in enumerate(enc, 1):
        c = ws.cell(5, ci, value=str(col).upper() if col else "")
        c.font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde()

    # Datos
    for fi, fila in enumerate(filas, 6):
        ws.row_dimensions[fi].height = 22
        for ci, col in enumerate(enc, 1):
            c = ws.cell(fi, ci, value=fila.get(col, ""))
            c.font = Font(name="Calibri", size=12)
            c.fill = PatternFill("solid", fgColor=COLOR_ROW_ALT if fi % 2 == 0 else "FFFFFF")
            c.alignment = Alignment(vertical="center")
            c.border = borde()

    ws.auto_filter.ref = f"A5:{ul}{5 + len(filas)}"
    ws.freeze_panes = "A6"
    autoajustar(ws)
    return nombre_hoja

def crear_menu(wb, escuelas_hojas, total, nombre_menu):
    ws = wb.create_sheet(title=nombre_menu, index=0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 4

    for f in [1, 2, 3, 4, 5]:
        ws.row_dimensions[f].height = 20 if f != 3 else 40
    for rng in ["B1:C1","B2:C2","B3:C3","B4:C4","B5:C5"]:
        ws.merge_cells(rng)

    t = ws["B3"]
    t.value = "INSCRITOS CULTURA 2026 — I"
    t.font = Font(name="Calibri", bold=True, size=20, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=COLOR_MENU_TITLE)
    t.alignment = Alignment(horizontal="center", vertical="center")

    s = ws["B4"]
    s.value = f"TALLER DE ORATORIA  |  {len(escuelas_hojas)} escuelas  |  {total} alumnos"
    s.font = Font(name="Calibri", size=12, color="B8CCE4")
    s.fill = PatternFill("solid", fgColor="243F60")
    s.alignment = Alignment(horizontal="center", vertical="center")

    for f in [1, 2, 5]:
        for col in ["B", "C"]:
            ws[f"{col}{f}"].fill = PatternFill("solid", fgColor=COLOR_MENU_TITLE)

    fa = 7
    ws.row_dimensions[fa].height = 22
    ws.merge_cells(f"B{fa}:C{fa}")
    i = ws[f"B{fa}"]
    i.value = "Haz clic en el nombre de una escuela para ir a su lista de alumnos:"
    i.font = Font(name="Calibri", italic=True, size=12, color="404040")
    i.alignment = Alignment(horizontal="left", vertical="center")
    fa += 1

    ws.row_dimensions[fa].height = 28
    for col, txt in [("B","ESCUELA PROFESIONAL"),("C","N.º ALUMNOS")]:
        c = ws[f"{col}{fa}"]
        c.value = txt
        c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde()
    fa += 1

    for idx, (escuela_original, nombre_hoja, cantidad) in enumerate(escuelas_hojas):
        ws.row_dimensions[fa].height = 26
        bg = COLOR_LINK_BG if idx % 2 == 0 else "1F5E99"

        cl = ws[f"B{fa}"]
        cl.value = escuela_original           # texto bonito con tildes
        cl.hyperlink = f"#{nombre_hoja}!A1"  # enlace con nombre limpio
        cl.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF", underline="single")
        cl.fill = PatternFill("solid", fgColor=bg)
        cl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cl.border = borde("2E75B6")

        cc = ws[f"C{fa}"]
        cc.value = cantidad
        cc.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=bg)
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc.border = borde("2E75B6")
        fa += 1

    ws.row_dimensions[fa].height = 28
    for col, val in [("B","TOTAL ALUMNOS ORATORIA"),("C", total)]:
        c = ws[f"{col}{fa}"]
        c.value = val
        c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_MENU_TITLE)
        c.alignment = Alignment(horizontal="center" if col=="C" else "left",
                                vertical="center", indent=1 if col=="B" else 0)
        c.border = borde()

def main():
    archivo = "INSCRITOS_CULTURA_2026_-_I.xlsx"
    if not os.path.exists(archivo):
        archivo = "INSCRITOS CULTURA 2026 - I.xlsx"
    if not os.path.exists(archivo):
        print("No se encontró el archivo Excel.")
        return

    print(f"Leyendo: {archivo}")
    enc, filas = leer_datos(archivo)

    col_taller  = next((h for h in enc if h and "TALLER"  in str(h).upper()), None)
    col_escuela = next((h for h in enc if h and "ESCUELA" in str(h).upper()), None)

    if not col_taller or not col_escuela:
        print("Columnas detectadas:", enc)
        return

    grupos = filtrar_y_agrupar(filas, col_taller, col_escuela)
    total  = sum(len(v) for v in grupos.values())
    print(f"Alumnos de Oratoria: {total} en {len(grupos)} escuelas")

    wb_out = openpyxl.Workbook()
    if "Sheet" in wb_out.sheetnames:
        del wb_out["Sheet"]

    nombre_menu = "MENU_PRINCIPAL"

    escuelas_hojas = []
    for escuela in sorted(grupos.keys()):
        filas_escuela = grupos[escuela]
        nombre_hoja   = crear_hoja_escuela(wb_out, escuela, filas_escuela, enc, nombre_menu)
        escuelas_hojas.append((escuela, nombre_hoja, len(filas_escuela)))
        print(f"  Hoja: {nombre_hoja} ({len(filas_escuela)} alumnos)")

    crear_menu(wb_out, escuelas_hojas, total, nombre_menu)

    salida = "ORATORIA_2026_POR_ESCUELA.xlsx"
    wb_out.save(salida)
    print(f"\nArchivo guardado: {salida}")

if __name__ == "__main__":
    main()
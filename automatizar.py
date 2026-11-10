"""Limpia datos_ventas.xlsx y genera reporte_ventas.xlsx.
Este es el script que se muestra en el video 1."""
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook, Workbook

MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
         "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12}


def leer_fecha(valor):
    """Acepta 2026-01-05, 05/01/2026 y 5-ene."""
    if isinstance(valor, datetime):
        return valor.date()
    texto = str(valor).strip().lower()
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    try:  # 5-ene
        dia, mes = texto.split("-")
        return datetime(2026, MESES[mes], int(dia)).date()
    except (ValueError, KeyError):
        return None


def leer_numero(valor):
    """Convierte 'S/ 4.50' o '3 unidades' en número. Devuelve None si no puede."""
    if isinstance(valor, (int, float)):
        return float(valor)
    limpio = str(valor).lower().replace("s/", "").replace("unidades", "").strip()
    try:
        return float(limpio)
    except ValueError:
        return None


# 1. Leer
wb = load_workbook("datos_ventas.xlsx")
ws = wb["Ventas"]
filas = list(ws.iter_rows(min_row=2, values_only=True))
print(f"Filas leídas: {len(filas)}")

# 2. Limpiar
limpias, vistos = [], set()
descartadas = 0
for fecha, cliente, producto, cantidad, precio in filas:
    f = leer_fecha(fecha)
    c = leer_numero(cantidad)
    p = leer_numero(precio)
    if f is None or c is None or p is None or not producto:
        descartadas += 1
        continue
    clave = (f, cliente, producto, c, p)
    if clave in vistos:          # duplicado
        descartadas += 1
        continue
    vistos.add(clave)
    limpias.append((f, cliente or "SIN NOMBRE", producto, int(c), p, int(c) * p))

print(f"Filas limpias: {len(limpias)} | descartadas: {descartadas}")

# 3. Agrupar
por_mes = defaultdict(float)
por_producto = defaultdict(lambda: [0, 0.0])
for f, cliente, producto, c, p, total in limpias:
    por_mes[f.strftime("%Y-%m")] += total
    por_producto[producto][0] += c
    por_producto[producto][1] += total

# 4. Escribir reporte
out = Workbook()
r = out.active
r.title = "Resumen mensual"
r.append(["Mes", "Ventas (S/)"])
for mes in sorted(por_mes):
    r.append([mes, round(por_mes[mes], 2)])

t = out.create_sheet("Top productos")
t.append(["Producto", "Unidades", "Ventas (S/)"])
for prod, (uni, tot) in sorted(por_producto.items(), key=lambda x: -x[1][1]):
    t.append([prod, uni, round(tot, 2)])

d = out.create_sheet("Datos limpios")
d.append(["Fecha", "Cliente", "Producto", "Cantidad", "Precio", "Total"])
for fila in limpias:
    d.append([fila[0], fila[1], fila[2], fila[3], fila[4], round(fila[5], 2)])

out.save("reporte_ventas.xlsx")

print(f"Total vendido: S/ {sum(por_mes.values()):,.2f}")
print("reporte_ventas.xlsx creado con 3 hojas")

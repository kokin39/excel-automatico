"""Genera un Excel 'desordenado' para el video 1 (datos de una bodega)."""
import random
from datetime import date, timedelta
from openpyxl import Workbook

random.seed(7)

PRODUCTOS = {
    "Arroz 1kg": 4.50, "Aceite 1L": 8.90, "Azúcar 1kg": 4.20,
    "Leche 400g": 4.80, "Fideos 500g": 3.50, "Atún": 6.50,
    "Café 250g": 12.00, "Galletas": 2.50,
}
CLIENTES = ["Bodega Ana", "Minimarket Luz", "Tienda Rosa", "Bodega El Sol",
            "Mercado Central", "Tienda José", "Bodega San Martín", "Minimarket Perú"]

FILAS = 400
inicio = date(2026, 1, 1)

filas = []
for _ in range(FILAS):
    d = inicio + timedelta(days=random.randint(0, 180))
    # Tres formatos de fecha distintos (el problema del video)
    formato = random.choice(["iso", "dmy", "texto"])
    if formato == "iso":
        fecha = d.strftime("%Y-%m-%d")
    elif formato == "dmy":
        fecha = d.strftime("%d/%m/%Y")
    else:
        meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
        fecha = f"{d.day}-{meses[d.month - 1]}"

    producto = random.choice(list(PRODUCTOS))
    precio = PRODUCTOS[producto]
    cliente = random.choice(CLIENTES + ["", ""])  # algunos vacíos

    qty = random.randint(1, 10)
    if random.random() < 0.08:
        cantidad = f"{qty} unidades"
    elif random.random() < 0.04:
        cantidad = ""
    else:
        cantidad = qty

    if random.random() < 0.15:
        precio_final = f"S/ {precio:.2f}"
    else:
        precio_final = round(precio * random.choice([1, 1, 1, 1.1]), 2)

    filas.append([fecha, cliente, producto, cantidad, precio_final])

# Duplicados (otro problema clásico)
for _ in range(18):
    filas.append(random.choice(filas).copy())

wb = Workbook()
ws = wb.active
ws.title = "Ventas"
ws.append(["Fecha", "Cliente", "Producto", "Cantidad", "Precio"])
for f in filas:
    ws.append(f)

wb.save("datos_ventas.xlsx")
print(f"datos_ventas.xlsx creado con {len(filas)} filas (desordenadas a propósito)")

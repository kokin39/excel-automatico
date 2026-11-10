# Automatiza tu Excel: de 2 horas a 2 minutos

Script en Python que limpia un Excel de ventas desordenado y genera un reporte
por mes y por producto.

🎬 Video explicativo: https://youtu.be/DTgSXWmmZ0w

## Qué hace

- Lee `datos_ventas.xlsx`.
- Acepta fechas en 3 formatos: `2026-01-05`, `05/01/2026` y `5-ene`.
- Convierte precios como `S/ 4.50` y cantidades como `3 unidades`.
- Elimina filas duplicadas y descarta las incompletas.
- Genera `reporte_ventas.xlsx` con 3 hojas:
  - **Resumen mensual**: ventas por mes.
  - **Top productos**: unidades y ventas por producto.
  - **Datos limpios**: la tabla ya normalizada.

## Requisitos

- Python 3.8 o superior.
- `openpyxl`:

```bash
pip install openpyxl
```

## Uso

```bash
python3 automatizar.py
```

Salida esperada:

```
Filas leídas: 418
Filas limpias: 384 | descartadas: 34
Total vendido: S/ 12,526.67
reporte_ventas.xlsx creado con 3 hojas
```

## Probar con datos nuevos

```bash
python3 crear_datos.py   # genera otro datos_ventas.xlsx desordenado
python3 automatizar.py
```

## Adaptarlo a tu Excel

En `automatizar.py` cambia:

1. El nombre del archivo: `datos_ventas.xlsx`.
2. El orden de las columnas en el bucle:
   `for fecha, cliente, producto, cantidad, precio in filas:`
3. Los formatos de fecha en `leer_fecha` si usas otro distinto.

## Contacto

¿Quieres que lo adapte a tu negocio? Escríbeme: kokin3839@gmail.com

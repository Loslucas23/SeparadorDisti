# Separar backup por distribuidor

Automatización en Python que toma un backup de ventas en Excel (`.xlsb` / `.xlsx`) y genera **un Excel por distribuidor**, con una hoja por año fiscal (FY24, FY25, …).

## Qué hace

- Lee todos los `.xlsb` / `.xlsx` de la carpeta principal.
- Cada hoja es un año fiscal (`FYxx`). Si no, el FY se calcula por la fecha (julio a junio).
- Ignora con aviso las hojas sin las columnas `DISTRIBUIDOR`, `Fecha` y `NETO s/IMP` (resúmenes, tablas dinámicas).
- **Combina archivos de forma inteligente**: para cada mes + distribuidor usa el archivo más reciente. Un archivo incremental agrega meses sin borrar los anteriores, y una corrección reemplaza solo ese mes.
- **Caché en Parquet**: los archivos que no cambiaron no se vuelven a leer.
- Escribe en paralelo (multiprocessing).
- Si un Excel de salida está abierto, lo saltea y avisa.
- Al final deja `control.txt` (totales de origen vs. generados) y agrega la corrida a `log.txt`.

## Uso

1. Copiar el/los backups en la carpeta principal.
2. Ejecutar `SEPARAR POR DISTRIBUIDOR.bat`. Instala las librerías la primera vez.
3. Elegir: `1` = todos los distribuidores, `2` = uno solo (por nombre o parte del nombre).

O directo con Python:

```bash
pip install pandas pyarrow pyxlsb xlsxwriter openpyxl
python Automatizacion/separar_por_distribuidor.py            # todos
python Automatizacion/separar_por_distribuidor.py "nombre"   # uno
```

Salida: `Por Distribuidor/<DISTRIBUIDOR>.xlsx`

Códigos de salida: `0` = OK · `2` = terminó con avisos (ver `control.txt`) · `1` = error

## Estructura

```
├── SEPARAR POR DISTRIBUIDOR.bat
└── Automatizacion/
    └── separar_por_distribuidor.py
```

> Los datos de ventas no forman parte de este repositorio.

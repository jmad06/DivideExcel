# DivideExcel

Divide un archivo Excel (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) en un CSV independiente por cada hoja, arrastrando el archivo sobre un `.bat`.

## Uso

1. Arrastra uno o varios archivos Excel sobre `DivideExcel.bat`.
2. Por cada archivo se crea una carpeta `<archivo>_csv` junto al original, con un CSV por hoja.

También se puede ejecutar directamente desde la terminal:

```bash
python DivideExcel.py archivo1.xlsx archivo2.xlsm
```

## Requisitos

- Python 3.9+
- [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip install openpyxl
```

## Configuración

Los ajustes se leen de [`config.json`](config.json), en la misma carpeta que el script. Si el archivo no existe, se genera automáticamente con los valores por defecto.

| Clave | Descripción | Por defecto |
|---|---|---|
| `delimitador` | Separador de campos del CSV. `";"` para abrir directo en Excel (España); `","` para pandas / Power BI. | `";"` |
| `codificacion` | Codificación del CSV. `"utf-8-sig"` conserva tildes y ñ al abrir en Excel. | `"utf-8-sig"` |
| `saltar_hojas_vacias` | Si es `true`, las hojas sin datos no generan CSV. | `true` |
| `extensiones` | Extensiones de Excel aceptadas como entrada. | `[".xlsx", ".xlsm", ".xltx", ".xltm"]` |

## Comportamiento

- Las fechas, horas y booleanos se formatean a texto legible en el CSV (`AAAA-MM-DD`, `HH:MM:SS`, `TRUE`/`FALSE`).
- Los nombres de hoja se convierten en nombres de archivo válidos en Windows; si dos hojas generan el mismo nombre, se añade un sufijo numérico.
- Si una hoja falla al procesarse, se omite con un aviso y el script continúa con el resto.
- Si una celda tiene fórmula pero no tiene valor calculado en caché (el libro no se guardó recalculado), se avisa por consola y la celda se exporta vacía.

## Licencia

Uso personal / interno.

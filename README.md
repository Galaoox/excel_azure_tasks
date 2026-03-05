# 🚀 Excel -> Azure Tasks CSV (Streamlit)

Aplicación web para transformar archivos Excel (`.xlsx`) en un CSV compatible con carga masiva de Azure Boards.

El objetivo es reducir trabajo manual: seleccionar hojas, mapear columnas, normalizar tareas y exportar un CSV consolidado listo para importar.

## 🛠️ Requisitos

- `uv` instalado
- Acceso a terminal (PowerShell, cmd o bash)

## 📦 Instalación

```bash
uv sync
```

## ▶️ Ejecutar la aplicación

```bash
uv run streamlit run app.py
```

## ✅ Ejecutar pruebas

```bash
uv run pytest -q
```

## 🧭 Flujo de uso (paso a paso)

1. Cargar un archivo `.xlsx`.
2. Seleccionar las hojas a trabajar.
3. Configurar cada hoja (en tabs):
   - fila de encabezado,
   - modo de prefijo (`sin prefijo`, `nombre hoja`, `custom`),
   - mapeo de columnas Excel -> Azure,
   - opción para omitir filas resumen (`total/subtotal/resumen`).
4. En la sección de construcción:
   - marcar qué hojas incluir,
   - (opcional) copiar mapeo de una hoja a otras.
5. Construir la tabla de tareas consolidada.
6. Revisar/editar en el editor:
   - búsqueda por título,
   - cambio de tipo de tarea individual o masivo.
7. Descargar CSV final para Azure.

## 📄 Formato de salida Azure

El archivo exportado usa columnas fijas:

| Columna Azure | Descripción |
|---|---|
| `Work Item Type` | Tipo de elemento (Task, Bug, etc.) |
| `Title` | Título final de la tarea |
| `Description` | Descripción |
| `Original Estimate` | Estimación inicial |
| `Remaining Work` | Trabajo restante |
| `Activity` | Actividad (manual o inferida) |

## 📌 Reglas de negocio importantes

- Se elimina cualquier fila sin título.
- Si `exclude summary rows` está activo (por defecto), se omiten filas cuyo título sea resumen:
  - `total`
  - `subtotal`
  - `resumen`
- Se usa normalización de mayúsculas/minúsculas y acentos para detectar esos casos.
- Si hay prefijo configurado, el título se construye como:
  - `{prefijo} - {titulo}`
- `Activity` se infiere cuando no viene informada.
- Si faltan columnas al copiar mapeo entre hojas, se avisa y esa hoja se omite de la copia.

## 🧪 Ejemplo corto

Entrada (conceptual):

- Hoja `Modulo A`
- Columna título contiene:
  - `Crear endpoint`
  - `Pruebas unitarias`
  - `Total`
- Prefijo configurado: `Modulo A`
- Omitir resumen: activado

Salida esperada:

- `Modulo A - Crear endpoint`
- `Modulo A - Pruebas unitarias`
- `Modulo A - Total` **no** se exporta

## 🛟 Troubleshooting

### Solo aparece una hoja en el resultado

- Verifica checkboxes de inclusión en la sección de construcción.
- Revisa el resumen de construcción por hoja (cuántas tareas aportó cada una).
- Si una hoja queda en 0, valida mapeo de título y filas resumen.

### Una hoja queda con 0 tareas

- Confirma que la columna mapeada como título no esté vacía.
- Si tus títulos son de resumen, desactiva temporalmente la exclusión de resumen para validar.

### Error al copiar mapeo entre hojas

- Si faltan columnas en hoja destino, se mostrará warning y no se aplicará el mapeo en esa hoja.

### Caracteres especiales/acentos en CSV

- El CSV se exporta en UTF-8 con BOM para mejorar compatibilidad con Excel.

## 🗂️ Estructura del proyecto

- `app.py`: interfaz Streamlit y orquestación del flujo.
- `src/domain`: modelos de dominio.
- `src/application`: reglas de negocio y transformación.
- `src/infrastructure`: lectura de Excel, escritura CSV y perfiles.
- `tests`: pruebas unitarias.

## ℹ️ Notas

- `script.py` se mantiene como referencia histórica; el flujo principal recomendado es la app Streamlit.

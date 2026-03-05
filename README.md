# Excel Azure Task CSV (Streamlit)

Aplicación web para convertir `.xlsx` en un CSV compatible con carga masiva de tareas en Azure.

## Requisitos

- `uv` instalado

## Instalación

```bash
uv sync
```

## Ejecutar app

```bash
uv run streamlit run app.py
```

## Ejecutar tests

```bash
uv run pytest -q
```

## Funcionalidades incluidas

- Selección de hojas a usar/omitir.
- Configuración por hoja de fila de encabezado (default fila 1).
- Mapeo de columnas Excel -> campos Azure.
- Edición de tipo de tarea individual y masiva sobre filas filtradas.
- Buscador de tareas por título.
- Prefijo por hoja (`sin prefijo`, `nombre hoja`, `custom`) con formato `{prefijo} - {titulo}`.
- Guardar/cargar perfiles JSON de configuración.
- Export CSV consolidado con columnas Azure fijas:
  - `Work Item Type`
  - `Title`
  - `Description`
  - `Original Estimate`
  - `Remaining Work`
  - `Activity`

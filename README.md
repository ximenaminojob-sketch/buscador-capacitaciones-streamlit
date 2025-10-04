# Buscador de Capacitaciones (Streamlit)

App para consultar **solo las capacitaciones realizadas** (con **fecha**) por **DNI**.

## Estructura esperada del Excel
- Hoja: `TECHINT`
- DNI en **columna C** desde fila **7** (`C7:C`).
- Encabezados de temas en **fila 6**, desde **G6** hacia la derecha.
- Matriz de valores desde **G7** (fechas/estados).

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py

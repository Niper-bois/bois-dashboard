# BOIS Dashboard V5

Paquete listo para publicar el Excel `BOIS_Excel_Master_V5.xlsx` como dashboard web con link.

## Qué incluye
- `app.py` → aplicación Streamlit lista para desplegar.
- `data/BOIS_Excel_Master_V5.xlsx` → backend inicial.
- `requirements.txt` → dependencias.
- `.streamlit/config.toml` → tema y ajustes.
- `render.yaml` → despliegue rápido en Render.

## Qué hace
- Resumen ejecutivo con KPIs y gráficos.
- Base de clientes filtrable.
- Vista financiera con payback, ROI y mejora EBITDA.
- Backlog de problemas y acciones.
- Matriz de módulos M01–M20.
- Informe operativo por cliente.
- Explorador del Excel para auditoría.
- Opción de sustituir el Excel desde la interfaz para refrescar el dashboard.

## Opción 1 — Streamlit Community Cloud (la más simple)
1. Crea un repositorio en GitHub.
2. Sube **todo** el contenido de esta carpeta al repositorio.
3. En Streamlit Community Cloud, crea una nueva app.
4. Selecciona tu repositorio, rama `main` y archivo `app.py`.
5. En "Advanced settings", usa Python 3.12.
6. Despliega y comparte la URL `*.streamlit.app`.

## Opción 2 — Render
1. Crea un repositorio en GitHub, GitLab o Bitbucket.
2. Sube **todo** el contenido de esta carpeta.
3. En Render, crea un nuevo Web Service y conecta ese repo.
4. Render detectará `render.yaml`.
5. Publica y comparte la URL `*.onrender.com`.

## Cambio de datos
- Opción rápida: dentro de la app, usa **"Sustituir Excel backend"**.
- Opción estable: reemplaza el archivo `data/BOIS_Excel_Master_V5.xlsx` en el repositorio y vuelve a desplegar.

## Nota importante
La app lee los **valores calculados** que guarda Excel. Si subes un archivo nuevo, conviene abrirlo, recalcular y guardar antes, para que las fórmulas lleguen actualizadas.

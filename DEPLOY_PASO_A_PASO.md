# Cómo ponerlo online y sacar el link

## Ruta recomendada: Streamlit Community Cloud

### Qué haces
1. Descomprime este paquete.
2. Crea un repo nuevo en GitHub, por ejemplo: `bois-dashboard-v5`.
3. Sube todos los archivos de la carpeta al repo.
4. Entra en Streamlit Community Cloud.
5. Pulsa **Create app**.
6. Elige:
   - **Repository**: tu repo
   - **Branch**: `main`
   - **Main file path**: `app.py`
7. En **Advanced settings**, selecciona **Python 3.12**.
8. En **App URL**, si quieres, pones el nombre.
9. Pulsa **Deploy**.
10. Te devuelve un link tipo `https://tu-nombre.streamlit.app`.

### Después
- Si quieres que cualquier persona entre: deja la app en público.
- Si quieres controlar acceso: ponla en privado e invita por email.

## Ruta alternativa: Render
1. Descomprime este paquete.
2. Súbelo a un repo GitHub/GitLab/Bitbucket.
3. Entra en Render.
4. Crea un **Web Service**.
5. Conecta el repo.
6. Render leerá `render.yaml`.
7. Despliega.
8. Te devuelve una URL `https://tu-app.onrender.com`.

## Qué no tienes que tocar
- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `render.yaml`
- `data/BOIS_Excel_Master_V5.xlsx`

## Cómo cambiar el Excel más adelante
### Opción 1
Desde la propia app, subes un nuevo Excel con el botón lateral.

### Opción 2
Reemplazas el archivo dentro de `data/` en el repo y vuelves a desplegar.

# 📊 Procesador de Estados Financieros - Reexpresión Monetaria (-6 Ceros)

Aplicación web desarrollada en **Python** y **Streamlit** para extraer balances y estados financieros desde archivos PDF o imágenes, reexpresar cifras eliminando los 6 ceros (dividiendo automáticamente entre `1,000,000` o un factor personalizable) y exportar el resultado a **Excel (.xlsx)** con formato contable profesional (encabezados azules, bordes finos, fórmulas y formato numérico `#,##0.00`).

Diseñada especialmente para desplegarse gratis y funcionar **24/7 en Streamlit Community Cloud**.

---

## 🚀 Guía de Despliegue en Streamlit Community Cloud (Gratis y Permanente)

Sigue estos 3 sencillos pasos para tener tu aplicación en línea con una URL pública accesible desde cualquier computadora o teléfono móvil, sin depender de tu máquina local ni de Antigravity.

---

### PASO 1: Subir el proyecto a GitHub

1. Ingresa a tu cuenta en [GitHub](https://github.com/) (o regístrate gratis si aún no tienes cuenta).
2. Haz clic en el botón verde **"New"** (Nuevo repositorio) en la esquina superior derecha o en la barra lateral.
3. Asigna un nombre al repositorio, por ejemplo: `reexpresion-financiera`.
4. Elige que sea **Public** (Público) para que Streamlit Community Cloud pueda acceder a él gratuitamente.
5. Haz clic en **"Create repository"**.
6. En la pantalla que aparece, haz clic en el enlace **"uploading an existing file"** (subir archivos existentes).
7. Arrastra y sube los 3 archivos de este proyecto:
   - `app.py`
   - `requirements.txt`
   - `README.md`
8. Haz clic en el botón verde **"Commit changes"** para guardar los archivos en la rama principal (`main`).

---

### PASO 2: Desplegar en Streamlit Community Cloud

1. Ingresa a [https://share.streamlit.io](https://share.streamlit.io) e inicia sesión con tu cuenta de **GitHub**.
2. Una vez dentro de tu panel de Streamlit, haz clic en el botón **"Create app"** o **"New app"**.
3. Selecciona la opción **"I already have an app"** (Ya tengo una aplicación).
4. Configura los campos del formulario:
   - **Repository:** Selecciona tu repositorio (ejemplo: `tu-usuario/reexpresion-financiera`).
   - **Branch:** Déjalo en `main` (o `master`).
   - **Main file path:** Escribe `app.py`.
   - *(Opcional)* En **App URL**, puedes personalizar el subdominio si está disponible (ejemplo: `mi-balance-reexpresado.streamlit.app`).
5. Haz clic en el botón azul **"Deploy!"**.
6. Streamlit Cloud leerá automáticamente `requirements.txt`, instalará los paquetes y levantará la aplicación en menos de 2 minutos.

---

### PASO 3: Compartir y Usar tu URL Permanente

1. En la esquina superior derecha de la app desplegada, haz clic en el botón de compartir o copia la URL directamente de la barra de direcciones de tu navegador (tendrá un formato similar a: `https://tu-proyecto.streamlit.app`).
2. **¡Listo!** Esta URL es permanente y pública:
   - Puedes abrirla desde cualquier navegador en PC, Mac, tablet o teléfono celular.
   - Funciona de forma 100% independiente: no requiere que tu ordenador esté encendido ni que la IDE/Antigravity esté abierta.
   - Cualquier cambio futuro que hagas en GitHub se actualizará automáticamente en la nube.

---

## 🛠️ Ejecución Local (Opcional)

Si deseas probar la aplicación localmente en tu equipo:

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar la aplicación
streamlit run app.py
```

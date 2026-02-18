# Sistema de Envío de Cartas — Modelo 347

Aplicación web local construida con **Streamlit** para el envío masivo y automatizado de cartas informativas del Modelo 347 (Declaración Anual de Operaciones con Terceras Personas) a través de Microsoft 365.

---

## Descripción del proyecto

El sistema permite:

1. **Cargar** un archivo ZIP con los PDFs de cada cliente (nombrados con la razón social).
2. **Cargar** un Excel con los datos de contacto (Nombre, Email, Dirección).
3. **Emparejar** automáticamente cada PDF con su cliente usando coincidencia difusa (`rapidfuzz`), con soporte para variaciones en la escritura de nombres y formas jurídicas (S.L., S.A., etc.).
4. **Revisar** los resultados del matching en una tabla interactiva y seleccionar los emails a enviar.
5. **Enviar** los emails de forma masiva mediante SMTP con Microsoft 365, adjuntando el PDF correspondiente.
6. **Descargar** un log de resultados en Excel con el estado de cada envío.

---

## Requisitos previos

- **Python 3.9** o superior
- Cuenta de **Microsoft 365** con SMTP habilitado
- **App Password** generada en la cuenta de Microsoft 365 (ver sección al final)
- Los PDFs deben estar **dentro de un archivo ZIP** y nombrados con la razón social exacta o aproximada del cliente
- El Excel debe contener las columnas: `Nombre`, `Email`, `Dirección`

---

## Instalación

### 1. Clonar o descargar el repositorio

```bash
git clone <url-del-repositorio>
cd envios-masivos
```

### 2. Crear un entorno virtual (recomendado)

```bash
python -m venv venv

# En Windows:
venv\Scripts\activate

# En macOS/Linux:
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

> **Nota:** `zipfile` es parte de la biblioteca estándar de Python y **no** se incluye en `requirements.txt`.

---

## Cómo ejecutar

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en el navegador en `http://localhost:8501`.

---

## Guía de uso

### Paso 1 — Preparar los archivos

**ZIP con PDFs:**
- Comprima todos los PDFs en un archivo `.ZIP`.
- Cada PDF debe llamarse con la razón social del cliente (p. ej. `Empresa García S.L..pdf`).
- Si tiene un `.RAR`, conviértalo a `.ZIP` usando 7-Zip, WinRAR u otra herramienta.

**Excel de clientes:**
- El archivo `.xlsx` debe contener obligatoriamente las columnas: `Nombre`, `Email`, `Dirección`.
- Las filas sin email serán descartadas automáticamente.

### Paso 2 — Configurar el SMTP (panel lateral)

| Campo           | Valor por defecto     | Descripción                        |
|-----------------|-----------------------|------------------------------------|
| Host SMTP       | `smtp.office365.com`  | Servidor de correo de Microsoft    |
| Puerto          | `587`                 | Puerto STARTTLS de Office 365      |
| Email remitente | —                     | Su dirección de correo empresarial |
| App Password    | —                     | Contraseña de aplicación (ver abajo) |

También puede personalizar el **asunto** y el **cuerpo** del email en el panel lateral.

### Paso 3 — Cargar archivos y ejecutar el matching

1. Suba el ZIP en la sección "Archivo ZIP con PDFs".
2. Suba el Excel en la sección "Excel con datos de clientes".
3. Pulse **Ejecutar matching**.

El sistema comparará cada PDF con los clientes del Excel usando coincidencia difusa:
- Primero busca por la columna `Nombre` (umbral ≥ 80%).
- Si no hay coincidencia, intenta con la columna `Dirección`.
- Los PDFs sin coincidencia se muestran en rojo al final.

### Paso 4 — Revisar y seleccionar

- Revise la tabla de resultados. El porcentaje de coincidencia aparece en verde (≥ 90%) o naranja (80–89%).
- Use **Seleccionar todos** / **Deseleccionar todos** o los checkboxes individuales.

### Paso 5 — Enviar

1. Pulse **Iniciar envío**.
2. La barra de progreso muestra el estado en tiempo real (`Enviando X de Y`).
3. Use **Cancelar envío** para detener el proceso tras el envío actual.
4. Configure la pausa entre envíos (1–10 segundos) para respetar los límites de Microsoft 365.

### Paso 6 — Descargar el log

Al finalizar, pulse **Descargar log en Excel** para obtener un archivo con:

| Columna          | Descripción                          |
|------------------|--------------------------------------|
| Nombre Archivo   | Nombre del PDF enviado               |
| Email Destino    | Dirección de correo del destinatario |
| Estado           | `Enviado` o `Error`                  |
| Mensaje de Error | Detalle del error (si aplica)        |
| Timestamp        | Fecha y hora del intento de envío    |

---

## Cómo obtener una App Password en Microsoft 365

> Las **App Passwords** permiten que aplicaciones externas usen su cuenta sin exponer la contraseña principal, y son necesarias cuando la autenticación multifactor (MFA) está activada.

### Pasos:

1. Acceda a [https://myaccount.microsoft.com](https://myaccount.microsoft.com) con su cuenta de trabajo.
2. Vaya a **Seguridad** → **Métodos de verificación adicionales** (o **Información de seguridad**).
3. Seleccione **Agregar método** → **Contraseña de aplicación**.
4. Escriba un nombre descriptivo (p. ej. `Envíos Modelo 347`) y pulse **Siguiente**.
5. Copie la contraseña generada y péguela en el campo **App Password** de la aplicación.

> **Importante:** La App Password solo se muestra una vez. Guárdela en un lugar seguro.

### Si el administrador tiene deshabilitadas las App Passwords:

Contacte con el administrador de Microsoft 365 de su organización para que habilite el uso de contraseñas de aplicación en el portal de administración (`admin.microsoft.com` → **Configuración** → **Servicios** → **Autenticación multifactor**).

---

## Estructura del proyecto

```
envios-masivos/
├── app.py            # Aplicación principal Streamlit
├── requirements.txt  # Dependencias Python
└── README.md         # Este archivo
```

---

## Dependencias principales

| Librería    | Versión mínima | Uso                                      |
|-------------|----------------|------------------------------------------|
| streamlit   | 1.28.0         | Interfaz web local                       |
| pandas      | 2.0.0          | Lectura del Excel y generación del log   |
| rapidfuzz   | 3.0.0          | Coincidencia difusa de nombres           |
| openpyxl    | 3.1.0          | Lectura/escritura de archivos `.xlsx`    |

---

## Notas y limitaciones

- **Límite de Microsoft 365:** El plan estándar permite hasta 10.000 emails/día y 30 mensajes/minuto. Use la pausa entre envíos para evitar bloqueos.
- **Tamaño de adjuntos:** Office 365 limita los adjuntos a 25 MB por mensaje.
- **Archivos RAR:** No son compatibles. Convierta a ZIP antes de subir.
- La aplicación se ejecuta **en local** y no almacena ningún dato en servidores externos.

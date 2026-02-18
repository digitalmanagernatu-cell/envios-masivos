"""
Sistema de Envío de Cartas - Modelo 347
Aplicación Streamlit para gestión y envío masivo de declaraciones informativas.
"""

import io
import re
import time
import smtplib
import zipfile
import datetime
import traceback
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

# ---------------------------------------------------------------------------
# Configuración de la página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Envíos Masivos - Modelo 347",
    page_icon="📨",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Estado de sesión — inicialización
# ---------------------------------------------------------------------------
_STATE_DEFAULTS = {
    "pdf_files": {},           # {nombre_sin_ext: bytes}
    "df_excel": None,          # DataFrame con columnas Nombre, Email, Dirección
    "matches": [],             # Lista de dicts con resultados del matching
    "unmatched": [],           # Lista de nombres PDF sin coincidencia
    "send_log": [],            # Log de envíos [{...}, ...]
    "sending": False,          # Flag de envío en curso
    "cancel_requested": False, # Flag de cancelación
    "matched_done": False,     # Matching ya ejecutado
    "sel_gen": 0,              # Generación de checkboxes (cambia al seleccionar/deseleccionar todos)
}
for _k, _v in _STATE_DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ---------------------------------------------------------------------------
# Utilidades de normalización
# ---------------------------------------------------------------------------
_LEGAL_SUFFIXES = re.compile(
    r"\b(s\.l\.u\.|slu|s\.l\.|sl|s\.a\.|sa)\b", re.IGNORECASE
)
_PUNCTUATION = re.compile(r"[.,;:()\-_/\\]+")
_SPACES = re.compile(r"\s+")


def normalize(text: str) -> str:
    """Normaliza un nombre para comparación fuzzy."""
    text = text.lower()
    text = _LEGAL_SUFFIXES.sub("", text)
    text = _PUNCTUATION.sub(" ", text)
    text = _SPACES.sub(" ", text)
    return text.strip()


# ---------------------------------------------------------------------------
# Separación de PDF completo en cartas individuales
# ---------------------------------------------------------------------------
def split_pdf_by_cif(pdf_bytes: bytes, cif: str) -> dict:
    """
    Divide un PDF completo en PDFs individuales detectando el CIF como marcador
    de inicio de carta. Devuelve un dict {nombre_cliente: bytes_pdf}.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    indices_inicio = [i for i in range(len(doc)) if cif in doc[i].get_text()]

    if not indices_inicio:
        doc.close()
        return {}

    pdf_dict = {}
    total = len(indices_inicio)

    for idx, start_page in enumerate(indices_inicio):
        end_page = indices_inicio[idx + 1] - 1 if idx + 1 < total else len(doc) - 1

        # Extraer nombre del cliente de la primera página de la carta
        texto = doc[start_page].get_text()
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        nombre_cliente = ""

        for i, linea in enumerate(lineas):
            # Patrón A: nombre tras "ejercicio" + "es de:"
            if "ejercicio" in linea.lower() and "es de:" in linea.lower():
                if i + 1 < len(lineas) and lineas[i + 1].lower() != "euros":
                    nombre_cliente = lineas[i + 1]
                    break
            # Patrón B: nombre tras la línea con el CIF
            if cif in linea and not nombre_cliente:
                if i + 1 < len(lineas):
                    candidato = lineas[i + 1]
                    if "Muy Sr" not in candidato and "URANO" not in candidato:
                        nombre_cliente = candidato
                        break

        if not nombre_cliente:
            nombre_cliente = f"Cliente_{idx + 1:03d}"

        # Limpiar nombre para usarlo como clave
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_cliente).strip()[:80]
        if not nombre_limpio:
            nombre_limpio = f"Cliente_{idx + 1:03d}"

        nuevo_doc = fitz.open()
        nuevo_doc.insert_pdf(doc, from_page=start_page, to_page=end_page)
        buf = io.BytesIO()
        nuevo_doc.save(buf)
        pdf_dict[nombre_limpio] = buf.getvalue()
        nuevo_doc.close()

    doc.close()
    return pdf_dict


# ---------------------------------------------------------------------------
# Lógica de matching
# ---------------------------------------------------------------------------
def run_matching(pdf_files: dict, df: pd.DataFrame):
    """
    Empareja PDFs con filas del Excel.

    Returns
    -------
    matches   : list[dict]  — coincidencias encontradas
    unmatched : list[str]   — nombres PDF sin coincidencia
    """
    matches = []
    unmatched = []

    # Preparar listas normalizadas para búsqueda
    nombres_norm = [normalize(str(n)) for n in df["Nombre"]]
    dirs_norm = [normalize(str(d)) for d in df["Dirección"]]

    for pdf_name in pdf_files:
        query_norm = normalize(pdf_name)

        # --- Búsqueda por Nombre ---
        result_name = process.extractOne(
            query_norm,
            nombres_norm,
            scorer=fuzz.token_sort_ratio,
        )

        if result_name and result_name[1] >= 80:
            idx = result_name[2]
            matches.append({
                "pdf_name": pdf_name,
                "cliente": str(df.at[idx, "Nombre"]),
                "email": str(df.at[idx, "Email"]),
                "score": result_name[1],
                "matched_by": "Nombre",
                "selected": True,
                "row_idx": idx,
            })
            continue

        # --- Búsqueda por Dirección (fallback) ---
        result_dir = process.extractOne(
            query_norm,
            dirs_norm,
            scorer=fuzz.token_sort_ratio,
        )

        if result_dir and result_dir[1] >= 80:
            idx = result_dir[2]
            matches.append({
                "pdf_name": pdf_name,
                "cliente": str(df.at[idx, "Nombre"]),
                "email": str(df.at[idx, "Email"]),
                "score": result_dir[1],
                "matched_by": "Dirección",
                "selected": True,
                "row_idx": idx,
            })
        else:
            unmatched.append(pdf_name)

    return matches, unmatched


# ---------------------------------------------------------------------------
# Lógica de envío de email
# ---------------------------------------------------------------------------
def send_email(
    smtp_host: str,
    smtp_port: int,
    sender_email: str,
    app_password: str,
    recipient_email: str,
    subject: str,
    body: str,
    pdf_name: str,
    pdf_bytes: bytes,
) -> None:
    """Envía un email con el PDF adjunto usando STARTTLS."""
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Adjunto PDF
    safe_filename = f"{pdf_name}.pdf"
    part = MIMEApplication(pdf_bytes, _subtype="pdf", Name=safe_filename)
    part.add_header(
        "Content-Disposition",
        "attachment",
        filename=safe_filename,
    )
    msg.attach(part)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender_email, app_password)
        server.send_message(msg)


# ---------------------------------------------------------------------------
# Sidebar — configuración SMTP
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Configuración SMTP")
    smtp_host = st.text_input("Host SMTP", value="smtp.office365.com")
    smtp_port = st.number_input("Puerto", value=587, min_value=1, max_value=65535, step=1)
    sender_email = st.text_input("Email remitente", placeholder="tu@empresa.com")
    app_password = st.text_input("App Password", type="password")

    st.divider()
    st.header("📄 Separación de PDF")
    cif_separator = st.text_input(
        "CIF de cabecera (marcador de inicio de carta)",
        value="B73798340",
        help="El sistema detecta el inicio de cada carta buscando este texto en la página.",
    )

    st.divider()
    st.header("📧 Plantilla de Email")

    default_subject = "Modelo 347 - Declaración Anual de Operaciones con Terceras Personas"
    email_subject = st.text_input("Asunto", value=default_subject)

    default_body = (
        "Estimado/a cliente,\n\n"
        "Adjunto encontrará la carta informativa relativa a la declaración anual "
        "de operaciones con terceras personas (Modelo 347) correspondiente al "
        "ejercicio fiscal.\n\n"
        "Por favor, revise detenidamente la información contenida y no dude en "
        "ponerse en contacto con nosotros si tiene alguna consulta o discrepancia.\n\n"
        "Atentamente,\nNATU Laboratories"
    )
    email_body = st.text_area("Cuerpo del email", value=default_body, height=200)

    st.divider()
    st.header("⏱️ Throttling")
    throttle_seconds = st.slider(
        "Pausa entre envíos (segundos)",
        min_value=1,
        max_value=10,
        value=2,
        step=1,
    )


# ---------------------------------------------------------------------------
# Título principal
# ---------------------------------------------------------------------------
st.title("📨 Sistema de Envío de Cartas — Modelo 347")
st.markdown(
    "Cargue el archivo ZIP con los PDFs y el Excel con los datos de clientes "
    "para realizar el envío masivo de forma automática."
)

# ---------------------------------------------------------------------------
# Sección 1 — Carga de archivos
# ---------------------------------------------------------------------------
st.header("1. Carga de Archivos")
col_zip, col_xls = st.columns(2)

with col_zip:
    st.subheader("📁 ZIP con PDFs individuales")
    uploaded_zip = st.file_uploader(
        "Suba el ZIP que contiene los PDFs",
        type=["zip"],
        help="Solo se aceptan archivos .ZIP. Si tiene un .RAR, conviértalo primero.",
    )

    if uploaded_zip is not None:
        _zip_id = (uploaded_zip.name, uploaded_zip.size)
        if st.session_state.get("_zip_file_id") == _zip_id:
            uploaded_zip = None  # mismo archivo, no reprocesar
        else:
            st.session_state["_zip_file_id"] = _zip_id

    if uploaded_zip is not None:
        try:
            with zipfile.ZipFile(io.BytesIO(uploaded_zip.read())) as zf:
                pdf_dict = {}
                skipped = []
                for name in zf.namelist():
                    if name.lower().endswith(".pdf") and not name.startswith("__MACOSX"):
                        base = name.split("/")[-1]
                        pdf_key = base[:-4]
                        pdf_dict[pdf_key] = zf.read(name)
                    elif not name.endswith("/"):
                        skipped.append(name)
            st.session_state["pdf_files"] = pdf_dict
            st.session_state["matched_done"] = False
            st.success(f"ZIP cargado: **{len(pdf_dict)} PDFs** encontrados.")
            if skipped:
                st.info(f"Archivos ignorados (no son PDF): {', '.join(skipped[:10])}")
        except zipfile.BadZipFile:
            st.error(
                "El archivo no es un ZIP válido. "
                "Si tiene un .RAR, conviértalo a .ZIP primero (7-Zip o WinRAR)."
            )

    st.divider()
    st.subheader("📄 PDF completo (todas las cartas juntas)")
    uploaded_pdf = st.file_uploader(
        "Suba el PDF con todas las cartas seguidas",
        type=["pdf"],
        help="El sistema detectará el inicio de cada carta por el CIF configurado en el panel lateral y las separará automáticamente.",
    )

    if uploaded_pdf is not None:
        _pdf_id = (uploaded_pdf.name, uploaded_pdf.size)
        if st.session_state.get("_pdf_file_id") == _pdf_id:
            uploaded_pdf = None  # mismo archivo, no reprocesar
        else:
            st.session_state["_pdf_file_id"] = _pdf_id

    if uploaded_pdf is not None:
        with st.spinner("Separando cartas del PDF..."):
            try:
                pdf_dict = split_pdf_by_cif(uploaded_pdf.read(), cif_separator)
                if not pdf_dict:
                    st.error(
                        f"No se encontró el marcador **{cif_separator}** en el PDF. "
                        "Revise el CIF configurado en el panel lateral."
                    )
                else:
                    st.session_state["pdf_files"] = pdf_dict
                    st.session_state["matched_done"] = False
                    st.success(f"PDF separado: **{len(pdf_dict)} cartas** detectadas.")
            except Exception as exc:
                st.error(f"Error al procesar el PDF: {exc}")

with col_xls:
    st.subheader("📊 Excel con datos de clientes")
    uploaded_excel = st.file_uploader(
        "Suba el Excel (.xlsx)",
        type=["xlsx"],
        help="Debe contener las columnas: Nombre, Email, Dirección",
    )

    if uploaded_excel is not None:
        _file_id = (uploaded_excel.name, uploaded_excel.size)
        if st.session_state.get("_excel_file_id") != _file_id:
            st.session_state["_excel_file_id"] = _file_id
        else:
            uploaded_excel = None  # mismo archivo, no reprocesar

    if uploaded_excel is not None:
        try:
            df = pd.read_excel(uploaded_excel, engine="openpyxl")
            # Normalizar nombres de columna: quitar espacios y homogeneizar
            df.columns = df.columns.str.strip()
            # Mapa de aliases para tolerar variantes comunes
            _col_aliases = {
                "nombre": "Nombre",
                "email": "Email",
                "correo": "Email",
                "e-mail": "Email",
                "direccion": "Dirección",
                "dirección": "Dirección",
                "direccion": "Dirección",
            }
            df.rename(
                columns={c: _col_aliases[c.lower()] for c in df.columns if c.lower() in _col_aliases},
                inplace=True,
            )
            required_cols = {"Nombre", "Email", "Dirección"}
            missing = required_cols - set(df.columns)
            if missing:
                st.error(
                    f"El Excel no contiene las columnas obligatorias: **{', '.join(missing)}**. "
                    "Por favor, revise el archivo."
                )
            else:
                df = df[["Nombre", "Email", "Dirección"]].dropna(subset=["Email"]).reset_index(drop=True)
                st.session_state["df_excel"] = df
                st.session_state["matched_done"] = False
                st.success(f"Excel cargado: **{len(df)} filas** válidas.")
        except Exception as exc:
            st.error(f"Error al leer el Excel: {exc}")

# ---------------------------------------------------------------------------
# Sección 2 — Matching
# ---------------------------------------------------------------------------
st.divider()
st.header("2. Emparejar PDFs con Clientes")

can_match = (
    len(st.session_state["pdf_files"]) > 0
    and st.session_state["df_excel"] is not None
)

if st.button("🔍 Ejecutar matching", disabled=not can_match, type="primary"):
    with st.spinner("Calculando coincidencias..."):
        matches, unmatched = run_matching(
            st.session_state["pdf_files"],
            st.session_state["df_excel"],
        )
    st.session_state["matches"] = matches
    st.session_state["unmatched"] = unmatched
    st.session_state["matched_done"] = True

if not can_match:
    st.info("Cargue el ZIP y el Excel para habilitar el matching.")

# ---------------------------------------------------------------------------
# Sección 3 — Resultados del matching
# ---------------------------------------------------------------------------
if st.session_state["matched_done"]:
    matches = st.session_state["matches"]
    unmatched = st.session_state["unmatched"]

    st.divider()
    st.header("3. Resultados del Matching")

    # --- Métricas de resumen ---
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total PDFs", len(st.session_state["pdf_files"]))
    col_m2.metric("Con coincidencia", len(matches))
    col_m3.metric("Sin coincidencia", len(unmatched))

    # --- Selección global ---
    if matches:
        col_sel1, col_sel2 = st.columns([1, 5])
        with col_sel1:
            if st.button("✅ Seleccionar todos"):
                for m in st.session_state["matches"]:
                    m["selected"] = True
                st.session_state["sel_gen"] += 1
        with col_sel2:
            if st.button("⬜ Deseleccionar todos"):
                for m in st.session_state["matches"]:
                    m["selected"] = False
                st.session_state["sel_gen"] += 1

        st.subheader("✅ PDFs con coincidencia")

        # Cabeceras de la tabla
        hdr = st.columns([3, 3, 3, 1, 2, 1])
        for col, label in zip(hdr, [
            "Nombre Archivo PDF", "Cliente Encontrado", "Email",
            "Score %", "Coincidencia por", "Enviar",
        ]):
            col.markdown(f"**{label}**")
        st.divider()

        for i, match in enumerate(st.session_state["matches"]):
            cols = st.columns([3, 3, 3, 1, 2, 1])
            cols[0].write(match["pdf_name"])
            cols[1].write(match["cliente"])
            cols[2].write(match["email"])
            score_color = "green" if match["score"] >= 90 else "orange"
            cols[3].markdown(
                f"<span style='color:{score_color}'><b>{match['score']}%</b></span>",
                unsafe_allow_html=True,
            )
            cols[4].write(match["matched_by"])
            _gen = st.session_state["sel_gen"]
            checked = cols[5].checkbox(
                "",
                value=match["selected"],
                key=f"sel_{_gen}_{i}",
                label_visibility="collapsed",
            )
            st.session_state["matches"][i]["selected"] = checked

    # --- PDFs sin coincidencia ---
    if unmatched:
        st.subheader("🔴 Sin coincidencia")
        st.markdown(
            "Los siguientes PDFs **no pudieron emparejarse** con ningún cliente del Excel:"
        )
        for name in unmatched:
            st.markdown(
                f"<div style='background:#ffe0e0;padding:6px 12px;"
                f"border-radius:4px;margin:4px 0;color:#c00'>"
                f"❌ {name}.pdf</div>",
                unsafe_allow_html=True,
            )
        # Botón de descarga Excel con los no encontrados
        df_unmatched = pd.DataFrame({"Archivo PDF sin emparejar": [f"{n}.pdf" for n in unmatched]})
        excel_buf = io.BytesIO()
        df_unmatched.to_excel(excel_buf, index=False, engine="openpyxl")
        st.download_button(
            label="⬇️ Descargar listado de no encontrados (.xlsx)",
            data=excel_buf.getvalue(),
            file_name="no_encontrados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------------------------------------------------------------------------
    # Sección 4 — Envío masivo
    # ---------------------------------------------------------------------------
    st.divider()
    st.header("4. Envío Masivo")

    selected_matches = [m for m in st.session_state["matches"] if m["selected"]]
    st.info(
        f"**{len(selected_matches)}** emails seleccionados para enviar. "
        "Configure el SMTP en el panel lateral antes de continuar."
    )

    ready_to_send = (
        bool(sender_email)
        and bool(app_password)
        and len(selected_matches) > 0
        and not st.session_state["sending"]
    )

    col_send, col_cancel = st.columns([2, 1])
    send_btn = col_send.button(
        "📤 Iniciar envío",
        disabled=not ready_to_send,
        type="primary",
    )
    cancel_btn = col_cancel.button(
        "⏹️ Cancelar envío",
        disabled=not st.session_state["sending"],
    )

    if cancel_btn:
        st.session_state["cancel_requested"] = True
        st.warning("Cancelación solicitada. El proceso se detendrá tras el envío actual.")

    if send_btn:
        st.session_state["sending"] = True
        st.session_state["cancel_requested"] = False
        st.session_state["send_log"] = []

        progress_bar = st.progress(0)
        status_text = st.empty()
        total = len(selected_matches)
        send_log = []

        for idx, match in enumerate(selected_matches):
            if st.session_state["cancel_requested"]:
                status_text.warning(f"Envío cancelado por el usuario tras {idx} emails.")
                break

            pdf_name = match["pdf_name"]
            recipient = match["email"]
            status_text.info(f"Enviando {idx + 1} de {total}: {pdf_name} → {recipient}")

            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                pdf_bytes = st.session_state["pdf_files"][pdf_name]
                send_email(
                    smtp_host=smtp_host,
                    smtp_port=int(smtp_port),
                    sender_email=sender_email,
                    app_password=app_password,
                    recipient_email=recipient,
                    subject=email_subject,
                    body=email_body,
                    pdf_name=pdf_name,
                    pdf_bytes=pdf_bytes,
                )
                send_log.append({
                    "Nombre Archivo": pdf_name,
                    "Email Destino": recipient,
                    "Estado": "Enviado",
                    "Mensaje de Error": "",
                    "Timestamp": timestamp,
                })
            except Exception as exc:
                error_msg = traceback.format_exc()
                send_log.append({
                    "Nombre Archivo": pdf_name,
                    "Email Destino": recipient,
                    "Estado": "Error",
                    "Mensaje de Error": str(exc),
                    "Timestamp": timestamp,
                })

            progress_bar.progress((idx + 1) / total)

            # Throttling — no esperar tras el último envío
            if idx < total - 1 and not st.session_state["cancel_requested"]:
                time.sleep(throttle_seconds)

        st.session_state["send_log"] = send_log
        st.session_state["sending"] = False
        st.session_state["cancel_requested"] = False

        # Resumen final
        enviados = sum(1 for r in send_log if r["Estado"] == "Enviado")
        fallidos = sum(1 for r in send_log if r["Estado"] == "Error")
        st.success(f"Proceso finalizado: **{enviados} enviados** correctamente.")
        if fallidos:
            st.error(f"**{fallidos} emails fallaron**. Consulte el log para más detalles.")

    # ---------------------------------------------------------------------------
    # Sección 5 — Log de resultados
    # ---------------------------------------------------------------------------
    if st.session_state["send_log"]:
        st.divider()
        st.header("5. Log de Envíos")

        df_log = pd.DataFrame(st.session_state["send_log"])

        # Colorear filas por estado
        def highlight_estado(row):
            color = "#d4edda" if row["Estado"] == "Enviado" else "#f8d7da"
            return [f"background-color: {color}"] * len(row)

        st.dataframe(
            df_log.style.apply(highlight_estado, axis=1),
            use_container_width=True,
        )

        # Métricas del log
        enviados_log = len(df_log[df_log["Estado"] == "Enviado"])
        fallidos_log = len(df_log[df_log["Estado"] == "Error"])
        col_l1, col_l2 = st.columns(2)
        col_l1.metric("Enviados correctamente", enviados_log)
        col_l2.metric("Fallidos", fallidos_log)

        # Descarga del log en Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_log.to_excel(writer, index=False, sheet_name="Log Envíos")
        buffer.seek(0)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 Descargar log en Excel",
            data=buffer,
            file_name=f"log_envios_347_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("---")
st.caption(
    "Sistema de Envío de Cartas Modelo 347 · "
    "Desarrollado con Streamlit · "
    f"{datetime.date.today().year}"
)

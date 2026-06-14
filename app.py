"""
Herramienta de disponibilidad docente — BDA Analytics
Corre con: streamlit run app.py
"""

import streamlit as st
import io, os, re, tempfile, json, zipfile, subprocess, sys
from pathlib import Path
from collections import defaultdict

# Asegurar que anthropic esté disponible
try:
    import anthropic as _anthropic_check
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "anthropic", "-q"], check=False)


from extractor_disponibilidad import procesar_archivo, exportar_excel
from detector_conflictos import (
    leer_propuesta, detectar_conflictos_propuesta,
    leer_disponibilidad_excel, detectar_conflictos_disponibilidad,
    exportar_conflictos,
)

# ── Constantes ─────────────────────────────────────────────────────────────
DIAS  = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]
HORAS = ["7-8","8-9","9-10","10-11","11-12","12-13","13-14",
         "14-15","15-16","16-17","17-18","18-19","19-20","20-21","21-22"]

def hora_a_set(txt):
    m = re.match(r"(\d+)\s*[-–]\s*(\d+)", str(txt).strip())
    return set(range(int(m.group(1)), int(m.group(2)))) if m else set()

def horas_solapan(h1, h2):
    return bool(hora_a_set(h1) & hora_a_set(h2))

# ── Config ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Disponibilidad Docente — BDA", page_icon="📅", layout="wide")

# ── Control de acceso ───────────────────────────────────────────────────────
from datetime import date as _date

_EXPIRACION = _date(2026, 7, 31)
_CODIGO_VALIDO = st.secrets.get("ACCESS_CODE", "") if hasattr(st, "secrets") else ""

_ADMIN_CODE = st.secrets.get("ADMIN_CODE", "") if hasattr(st, "secrets") else ""

if _CODIGO_VALIDO or _ADMIN_CODE:  # solo activar si hay códigos configurados
    if not st.session_state.get("_acceso_ok"):
        st.title("📅 Disponibilidad Docente — BDA Analytics")
        st.markdown("#### Ingresa tu código de acceso")
        codigo_input = st.text_input("Código", type="password", placeholder="Escribe el código que te enviaron")
        if st.button("Ingresar", type="primary"):
            codigo = codigo_input.strip()
            if _ADMIN_CODE and codigo == _ADMIN_CODE:
                # Admin: acceso ilimitado, sin expiración
                st.session_state["_acceso_ok"] = True
                st.rerun()
            elif codigo == _CODIGO_VALIDO:
                if _date.today() > _EXPIRACION:
                    st.error("⏰ Tu período de acceso ha expirado. Contacta a BDA Analytics para renovar.")
                else:
                    st.session_state["_acceso_ok"] = True
                    st.rerun()
            else:
                st.error("Código incorrecto. Verifica con BDA Analytics.")
        st.stop()

st.markdown("""
<style>
  .block-container{padding-top:1.5rem}
  h1,h2,h3{color:#1A3557}
  .tag{display:inline-block;background:#E8F0FE;color:#1A3557;border-radius:12px;
       padding:2px 10px;font-size:0.8rem;margin:2px}
  .tag-curso{background:#FFF0C2;color:#7A5000}
  .tag-dia{background:#E6F4EA;color:#1E6E3A}
  .card{background:#F8F9FA;border-radius:10px;padding:1rem 1.2rem;margin-bottom:0.6rem;
        border-left:4px solid #1A3557;color:#1a1a1a}
  .card strong{color:#1A3557}
  .card small{color:#555 !important}
  .card-verde{border-left-color:#34A853}
  .chat-user{background:#E8F0FE;border-radius:12px;padding:0.7rem 1rem;margin:0.4rem 0;color:#1a1a1a}
  .chat-bot{background:#F1F3F4;border-radius:12px;padding:0.7rem 1rem;margin:0.4rem 0;color:#1a1a1a}
</style>
""", unsafe_allow_html=True)

# ── Helpers de presentación ────────────────────────────────────────────────
def disp_resumen(disp):
    partes = []
    for dia in DIAS:
        hs = disp.get(dia, [])
        if hs:
            partes.append(f"**{dia[:3]}**: {hs[0]}–{hs[-1].split('-')[1]}" if len(hs) > 1 else f"**{dia[:3]}**: {hs[0]}")
    return "  ·  ".join(partes) if partes else "Sin disponibilidad"

def disp_tabla(disp):
    filas = []
    for dia in DIAS:
        hs = disp.get(dia, [])
        if hs:
            filas.append(f"| {dia} | {', '.join(hs)} |")
    if not filas:
        return "Sin disponibilidad marcada"
    return "| Día | Horas disponibles |\n|---|---|\n" + "\n".join(filas)

def nombre_corto(nombre):
    """MERINO ESCUZA CAMILA → Camila Merino"""
    if not nombre:
        return "—"
    partes = nombre.strip().split()
    if len(partes) >= 3:
        return f"{partes[-1].capitalize()} {partes[0].capitalize()}"
    return nombre.title()

# ── Estado de sesión ───────────────────────────────────────────────────────
if "datos" not in st.session_state:
    st.session_state["datos"] = []          # lista de registros procesados
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("bda_logo.svg", use_container_width=True)
    st.markdown("### Cargar datos")

    word_files = st.file_uploader(
        "Word de disponibilidad (.docx) — selección múltiple",
        type=["docx"], accept_multiple_files=True,
    )
    zip_file = st.file_uploader(
        "O sube un ZIP con todos los Word (.zip)",
        type=["zip"],
    )
    propuesta_file = st.file_uploader(
        "Propuesta de horarios (.csv) — opcional",
        type=["csv"],
    )

    if st.button("⚡ Procesar archivos", type="primary", use_container_width=True):
        if not word_files and not zip_file:
            st.error("Sube al menos un archivo Word o un ZIP.")
        else:
            # Recopilar bytes de todos los docx
            archivos_a_procesar = []  # lista de (nombre, bytes)
            for f in word_files:
                archivos_a_procesar.append((f.name, f.getvalue()))
            if zip_file:
                with zipfile.ZipFile(io.BytesIO(zip_file.getvalue())) as zf:
                    for entry in zf.namelist():
                        if entry.lower().endswith(".docx") and not entry.startswith("__"):
                            archivos_a_procesar.append((Path(entry).name, zf.read(entry)))

            total = len(archivos_a_procesar)
            with st.spinner(f"Procesando {total} archivo(s)..."):
                registros = []
                progreso = st.progress(0)
                with tempfile.TemporaryDirectory() as tmp:
                    for i, (nombre_f, contenido) in enumerate(archivos_a_procesar):
                        ruta = Path(tmp) / nombre_f
                        ruta.write_bytes(contenido)
                        registros.append(procesar_archivo(str(ruta)))
                        progreso.progress((i + 1) / total)
                st.session_state["datos"] = registros
                st.session_state["chat_history"] = []

                # Conflictos si hay propuesta
                conflictos = []
                conflictos_bytes = None
                if propuesta_file:
                    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as tc:
                        tc.write(propuesta_file.getvalue())
                        tc_path = tc.name
                    filas_prop = leer_propuesta(tc_path)
                    conflictos = detectar_conflictos_propuesta(filas_prop)
                    os.unlink(tc_path)
                st.session_state["conflictos"] = conflictos

            exitosos = sum(1 for r in registros if not r.get("error"))
            st.success(f"✓ {exitosos} docente(s) procesado(s)")

    st.divider()
    st.markdown("### Chat con IA")
    # Leer key desde Streamlit Secrets si está configurada (para deploy en la nube)
    _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
    if _secret_key and not st.session_state.get("api_key"):
        st.session_state["api_key"] = _secret_key

    api_key = st.text_input("API Key de Anthropic (opcional)", type="password",
                            value="",
                            help="Para activar el chat en lenguaje natural. "
                                 "Si el administrador configuró la key en Secrets, no necesitas escribirla.")
    if api_key:
        st.session_state["api_key"] = api_key

    if st.session_state["datos"]:
        st.divider()
        # Descargar disponibilidad
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            exportar_excel(st.session_state["datos"], tmp.name)
            with open(tmp.name, "rb") as f:
                xls_bytes = f.read()
            os.unlink(tmp.name)
        st.download_button("📥 Disponibilidad (.xlsx)", xls_bytes,
                           "disponibilidad.xlsx", use_container_width=True)

        if st.session_state.get("conflictos"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                exportar_conflictos(st.session_state["conflictos"], tmp.name)
                with open(tmp.name, "rb") as f:
                    conf_bytes = f.read()
                os.unlink(tmp.name)
            st.download_button("📥 Conflictos (.xlsx)", conf_bytes,
                               "conflictos.xlsx", use_container_width=True)

# ── Main ───────────────────────────────────────────────────────────────────
st.title("📅 Disponibilidad Docente")
st.caption("BDA Analytics · Universidad de Lima")

datos = st.session_state["datos"]
exitosos = [r for r in datos if not r.get("error")]

if not exitosos:
    st.info("👈 Sube los archivos Word en el panel izquierdo y presiona **Procesar archivos**.")
    st.stop()

# Métricas rápidas
m1, m2, m3, m4 = st.columns(4)
todos_cursos = set()
for r in exitosos:
    todos_cursos.update(r["cursos_asignados"])
    todos_cursos.update(r["cursos_adicionales"])
m1.metric("Docentes cargados", len(exitosos))
m2.metric("Cursos en el sistema", len(todos_cursos))
m3.metric("Conflictos detectados", len(st.session_state.get("conflictos", [])))
dias_activos = sum(1 for r in exitosos for d in DIAS if r["disponibilidad"].get(d))
m4.metric("Disponibilidades registradas", dias_activos)

st.divider()

# ── Tabs ───────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "👤 Por profesor",
    "📚 Por curso",
    "🕐 Por horario",
    "💬 Preguntar",
])

# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — Por profesor
# ══════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("¿Qué puede enseñar un profesor y cuándo?")
    nombres_lista = sorted([r["nombre"] or r["archivo"] for r in exitosos])
    sel = st.selectbox("Selecciona un docente", ["— elige uno —"] + nombres_lista,
                       key="tab1_profesor")

    if sel != "— elige uno —":
        reg = next((r for r in exitosos if (r["nombre"] or r["archivo"]) == sel), None)
        if reg:
            c1, c2 = st.columns([1, 1])

            with c1:
                st.markdown("#### Disponibilidad horaria")
                st.markdown(disp_tabla(reg["disponibilidad"]))

            with c2:
                st.markdown("#### Cursos")
                if reg["cursos_asignados"]:
                    st.markdown("**Asignados:**")
                    for c in reg["cursos_asignados"]:
                        st.markdown(f'<span class="tag tag-curso">{c}</span>', unsafe_allow_html=True)
                if reg["cursos_adicionales"]:
                    st.markdown("**También puede dictar:**")
                    for c in reg["cursos_adicionales"]:
                        st.markdown(f'<span class="tag">{c}</span>', unsafe_allow_html=True)
                if reg["jps"]:
                    st.markdown("**JPs sugeridos:**")
                    for curso, jp in reg["jps"].items():
                        st.markdown(f"- {curso}: **{jp}**")
                if reg.get("ingles"):
                    st.markdown(f"🌐 Puede dictar en inglés: {reg['ingles']}")
                if reg.get("otras_instituciones"):
                    st.markdown(f"🏫 Dicta en otras instituciones: {reg['otras_instituciones']}")

# ══════════════════════════════════════════════════════════════════════════
# TAB 2 — Por curso
# ══════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("¿Quién puede enseñar este curso?")

    # Construir índice curso → docentes
    indice_cursos = defaultdict(list)
    for r in exitosos:
        for curso in r["cursos_asignados"]:
            indice_cursos[curso].append((r, "ASIGNADO"))
        for curso in r["cursos_adicionales"]:
            if r not in [x[0] for x in indice_cursos[curso]]:
                indice_cursos[curso].append((r, "ADICIONAL"))

    cursos_lista = sorted(indice_cursos.keys())
    c_busca = st.text_input("Buscar curso", placeholder="Ej: Storytelling, Marketing, Narrativa...")
    if c_busca:
        cursos_lista = [c for c in cursos_lista if c_busca.upper() in c.upper()]

    sel_curso = st.selectbox("Selecciona el curso", ["— elige uno —"] + cursos_lista, key="tab2_curso")

    if sel_curso != "— elige uno —":
        docentes_curso = indice_cursos[sel_curso]
        st.markdown(f"**{len(docentes_curso)} docente(s)** pueden dictar **{sel_curso}**:")

        for reg, tipo in docentes_curso:
            color = "card-verde" if tipo == "ASIGNADO" else "card"
            badge = "🟢 Asignado" if tipo == "ASIGNADO" else "🔵 Disponible"
            nombre_display = (reg["nombre"] or reg["archivo"]).title()
            dias_str = "  ·  ".join(
                f"{d}: {reg['disponibilidad'][d][0]}–{reg['disponibilidad'][d][-1].split('-')[1]}"
                for d in DIAS if reg["disponibilidad"].get(d)
            ) or "Sin disponibilidad registrada"
            st.markdown(f"""
            <div class="card {color}">
              <strong>{nombre_display}</strong> &nbsp; <small>{badge}</small><br>
              <small style="color:#555">{dias_str}</small>
            </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# TAB 3 — Por horario
# ══════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("¿Quién está disponible en este horario?")

    col_d, col_h = st.columns(2)
    with col_d:
        dia_sel = st.selectbox("Día", DIAS, key="tab3_dia")
    with col_h:
        hora_sel = st.selectbox("Franja horaria", HORAS, key="tab3_hora")

    hora_set = hora_a_set(hora_sel)

    disponibles = []
    no_disponibles = []
    for r in exitosos:
        horas_dia = set()
        for h in r["disponibilidad"].get(dia_sel, []):
            horas_dia.update(hora_a_set(h))
        if hora_set & horas_dia:
            disponibles.append(r)
        else:
            no_disponibles.append(r)

    st.markdown(f"**{len(disponibles)}** docente(s) disponibles el **{dia_sel} de {hora_sel}h**:")

    if disponibles:
        cols = st.columns(2)
        for i, r in enumerate(disponibles):
            nombre_display = (r["nombre"] or r["archivo"]).title()
            with cols[i % 2]:
                st.markdown(f"""
                <div class="card card-verde">
                  <strong>{nombre_display}</strong>
                </div>""", unsafe_allow_html=True)
    else:
        st.warning("Ningún docente registró disponibilidad en ese horario.")

    if no_disponibles:
        with st.expander(f"Ver los {len(no_disponibles)} docentes NO disponibles en ese horario"):
            for r in no_disponibles:
                st.markdown(f"- {r['nombre'].title()}")

# ══════════════════════════════════════════════════════════════════════════
# TAB 4 — Chat con IA
# ══════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Pregunta en lenguaje natural")

    api_key_activa = st.session_state.get("api_key", "")

    if not api_key_activa:
        st.info("Ingresa tu API Key de Anthropic en el panel izquierdo para activar el chat.")
    else:
        # Construir contexto de datos para el modelo
        def construir_contexto():
            lineas = ["DATOS DE DISPONIBILIDAD DOCENTE:\n"]
            for r in exitosos:
                lineas.append(f"DOCENTE: {r['nombre']} (código {r['codigo']})")
                for dia in DIAS:
                    hs = r["disponibilidad"].get(dia, [])
                    if hs:
                        lineas.append(f"  {dia}: {', '.join(hs)}")
                if r["cursos_asignados"]:
                    lineas.append(f"  CURSOS ASIGNADOS: {'; '.join(r['cursos_asignados'])}")
                if r["cursos_adicionales"]:
                    lineas.append(f"  PUEDE TAMBIÉN DICTAR: {'; '.join(r['cursos_adicionales'])}")
                if r["jps"]:
                    jps = "; ".join(f"{k}→{v}" for k, v in r["jps"].items())
                    lineas.append(f"  JPS SUGERIDOS: {jps}")
                lineas.append("")
            return "\n".join(lineas)

        SYSTEM_PROMPT = f"""Eres un asistente de planificación académica para la Universidad de Lima.
Tienes acceso a la disponibilidad horaria y cursos de los docentes de la Facultad de Comunicación.
Respondes preguntas sobre quién puede enseñar qué curso, en qué horarios están disponibles los docentes,
y qué opciones hay para cubrir una sección. Eres directo y concreto. Si no tienes la información,
lo dices claramente.

{construir_contexto()}"""

        # Historial de chat
        for msg in st.session_state["chat_history"]:
            if msg["role"] == "user":
                st.markdown(f'<div class="chat-user">👤 {msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="chat-bot">🤖 {msg["content"]}</div>', unsafe_allow_html=True)

        # Input
        with st.form("chat_form", clear_on_submit=True):
            pregunta = st.text_input(
                "Tu pregunta",
                placeholder="Ej: ¿Quién puede dar Storytelling los martes? ¿Qué horarios tiene libre García Contto?",
                label_visibility="collapsed",
            )
            enviado = st.form_submit_button("Enviar →", use_container_width=True)

        if enviado and pregunta.strip():
            st.session_state["chat_history"].append({"role": "user", "content": pregunta})

            with st.spinner("Consultando..."):
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key_activa)
                    messages = [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state["chat_history"]
                    ]
                    respuesta = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=1024,
                        system=SYSTEM_PROMPT,
                        messages=messages,
                    )
                    texto = respuesta.content[0].text
                except ImportError:
                    texto = "⚠️ Falta instalar la librería anthropic. Abre la Terminal, corre: `pip3 install anthropic` y reinicia la app."
                except Exception as e:
                    texto = f"⚠️ Error: {str(e)}"

            st.session_state["chat_history"].append({"role": "assistant", "content": texto})
            st.rerun()

        if st.session_state["chat_history"]:
            if st.button("🗑 Limpiar conversación", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

        # Preguntas de ejemplo
        st.divider()
        st.markdown("**Ejemplos de preguntas:**")
        ejemplos = [
            "¿Quién puede dictar Storytelling y en qué horarios está disponible?",
            "¿Qué cursos puede enseñar el docente Merino y cuándo tiene libre?",
            "¿Quién tiene disponibilidad los lunes de 9 a 13?",
            "¿Hay alguien que pueda cubrir Narrativa Audiovisual los martes?",
        ]
        for ej in ejemplos:
            st.markdown(f"- *{ej}*")

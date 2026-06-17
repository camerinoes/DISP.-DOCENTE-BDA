"""
Herramienta de disponibilidad docente — Universidad de Lima
Corre con: streamlit run app.py
"""

import streamlit as st
import io, os, re, tempfile, json, zipfile, subprocess, sys
from pathlib import Path
from collections import defaultdict

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

NARANJA = "#F05A22"   # naranja Ulima
NEGRO   = "#1A1A1A"
GRIS    = "#F5F5F5"

def hora_a_set(txt):
    m = re.match(r"(\d+)\s*[-–]\s*(\d+)", str(txt).strip())
    return set(range(int(m.group(1)), int(m.group(2)))) if m else set()

def horas_solapan(h1, h2):
    return bool(hora_a_set(h1) & hora_a_set(h2))

# ── Config ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Disponibilidad Docente — U de Lima", page_icon="📅", layout="wide")

# ── Control de acceso ───────────────────────────────────────────────────────
from datetime import date as _date

_EXPIRACION = _date(2026, 7, 31)
_CODIGO_VALIDO = st.secrets.get("ACCESS_CODE", "") if hasattr(st, "secrets") else ""
_ADMIN_CODE = st.secrets.get("ADMIN_CODE", "") if hasattr(st, "secrets") else ""

if _CODIGO_VALIDO or _ADMIN_CODE:
    if not st.session_state.get("_acceso_ok"):
        _, col_c, _ = st.columns([1, 2, 1])
        with col_c:
            st.markdown(f"""
            <div style="text-align:center;margin-bottom:1rem">
              <div style="font-size:2rem;font-weight:800;color:{NARANJA};letter-spacing:-1px">
                Disponibilidad Docente
              </div>
              <div style="font-size:0.9rem;color:#666;margin-top:0.3rem">
                Universidad de Lima
              </div>
            </div>
            """, unsafe_allow_html=True)
            st.divider()
            st.markdown("**Ingresa tu código de acceso**")
            codigo_input = st.text_input(
                "Código", type="password",
                placeholder="Escribe el código que te enviaron",
                label_visibility="collapsed",
            )
            if st.button("Ingresar →", type="primary", use_container_width=True):
                codigo = codigo_input.strip()
                if _ADMIN_CODE and codigo == _ADMIN_CODE:
                    st.session_state["_acceso_ok"] = True
                    st.rerun()
                elif codigo == _CODIGO_VALIDO:
                    if _date.today() > _EXPIRACION:
                        st.error("⏰ Tu período de acceso ha expirado.")
                    else:
                        st.session_state["_acceso_ok"] = True
                        st.rerun()
                else:
                    st.error("Código incorrecto.")
        st.stop()

# ── Estilos globales ────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  .block-container{{padding-top:1.5rem}}
  h1,h2,h3{{color:{NEGRO}}}

  /* Botones primarios → naranja */
  .stButton > button[kind="primary"] {{
    background-color:{NARANJA} !important;
    border-color:{NARANJA} !important;
    color:#fff !important;
  }}
  .stButton > button[kind="primary"]:hover {{
    background-color:#d94d1a !important;
    border-color:#d94d1a !important;
  }}

  /* Tabs: borde activo naranja */
  .stTabs [data-baseweb="tab-list"] {{
    border-bottom: 2px solid #E0E0E0;
    gap: 0.5rem;
  }}
  .stTabs [aria-selected="true"] {{
    border-bottom: 3px solid {NARANJA} !important;
    color: {NARANJA} !important;
    font-weight: 700;
  }}

  /* Pasos del sidebar */
  .step-label {{
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 10px;
    background: {NARANJA};
    color: #fff;
    border-radius: 8px;
    font-weight: 700;
    font-size: 0.85rem;
    margin: 0.6rem 0 0.4rem 0;
  }}

  /* Cards */
  .tag{{display:inline-block;background:#FFF0E8;color:{NARANJA};border-radius:12px;
       padding:2px 10px;font-size:0.8rem;margin:2px;border:1px solid #F9C0A8}}
  .tag-curso{{background:#FFF0E8;color:#C04400}}
  .card{{background:#FAFAFA;border-radius:10px;padding:1rem 1.2rem;margin-bottom:0.6rem;
        border-left:4px solid {NEGRO};color:{NEGRO}}}
  .card strong{{color:{NEGRO}}}
  .card small{{color:#666 !important}}
  .card-ok{{border-left-color:{NARANJA}}}

  /* Chat */
  .chat-user{{background:#FFF0E8;border-radius:12px;padding:0.7rem 1rem;margin:0.4rem 0;color:{NEGRO}}}
  .chat-bot{{background:#F5F5F5;border-radius:12px;padding:0.7rem 1rem;margin:0.4rem 0;color:{NEGRO}}}
</style>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────────────────
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
    if not nombre:
        return "—"
    partes = nombre.strip().split()
    if len(partes) >= 3:
        return f"{partes[-1].capitalize()} {partes[0].capitalize()}"
    return nombre.title()

# ── Estado de sesión ───────────────────────────────────────────────────────
if "datos" not in st.session_state:
    st.session_state["datos"] = []
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="font-size:1rem;font-weight:800;color:{NARANJA};
                letter-spacing:-0.5px;padding:0.5rem 0 0.2rem 0">
      📅 Disponibilidad Docente
    </div>
    <div style="font-size:0.75rem;color:#888;margin-bottom:0.5rem">
      Universidad de Lima
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    # ── Paso 1 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="step-label">1 · Cargar archivos</div>', unsafe_allow_html=True)

    word_files = st.file_uploader(
        "Word de disponibilidad (.docx)",
        type=["docx"], accept_multiple_files=True,
        help="Selecciona uno o varios formularios Word.",
    )
    zip_file = st.file_uploader(
        "O sube un ZIP con todos los Word",
        type=["zip"],
    )
    with st.expander("➕ Propuesta de horarios (opcional)"):
        propuesta_file = st.file_uploader(
            "Propuesta de horarios (.csv)",
            type=["csv"],
        )

    # ── Paso 2 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="step-label">2 · Procesar</div>', unsafe_allow_html=True)

    if st.button("⚡ Procesar archivos", type="primary", use_container_width=True):
        if not word_files and not zip_file:
            st.error("Sube al menos un archivo Word o ZIP.")
        else:
            archivos_a_procesar = []
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

                conflictos = []
                if propuesta_file:
                    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as tc:
                        tc.write(propuesta_file.getvalue())
                        tc_path = tc.name
                    filas_prop = leer_propuesta(tc_path)
                    conflictos = detectar_conflictos_propuesta(filas_prop)
                    os.unlink(tc_path)
                st.session_state["conflictos"] = conflictos

            exitosos_sb = sum(1 for r in registros if not r.get("error"))
            st.success(f"✓ {exitosos_sb} docente(s) procesado(s)")

    # ── Paso 3 ──────────────────────────────────────────────────────────────
    st.markdown('<div class="step-label">3 · Descargar resultados</div>', unsafe_allow_html=True)

    datos_ok = [r for r in st.session_state["datos"] if not r.get("error")]
    if datos_ok:
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
    else:
        st.caption("Disponible luego de procesar.")

    st.divider()

    # ── Chat con IA ─────────────────────────────────────────────────────────
    st.markdown(f'<div style="font-size:0.8rem;font-weight:700;color:{NEGRO}">🤖 Chat con IA</div>',
                unsafe_allow_html=True)

    _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
    if _secret_key and not st.session_state.get("api_key"):
        st.session_state["api_key"] = _secret_key

    if st.session_state.get("api_key"):
        st.success("✓ Chat con IA activado")
    else:
        api_key = st.text_input("API Key de Anthropic", type="password")
        if api_key:
            st.session_state["api_key"] = api_key

# ── Main ───────────────────────────────────────────────────────────────────
st.markdown(f'<div style="font-size:2rem;font-weight:800;color:{NARANJA};line-height:1.2;margin-bottom:0.1rem">Disponibilidad Docente</div>', unsafe_allow_html=True)
st.caption("Universidad de Lima")

datos = st.session_state["datos"]
exitosos = [r for r in datos if not r.get("error")]

if not exitosos:
    st.markdown(f"""
    <div style="background:{GRIS};border-radius:12px;padding:1.5rem 2rem;
                border-left:4px solid {NARANJA};margin:1rem 0">
      <div style="font-weight:700;font-size:1.05rem;color:{NEGRO};margin-bottom:0.5rem">
        👋 Bienvenido/a
      </div>
      <div style="color:#555;font-size:0.9rem">
        Sigue los pasos del panel izquierdo para comenzar:
        <ol style="margin:0.5rem 0 0 1rem;padding:0">
          <li>Carga los archivos Word (.docx) o un ZIP</li>
          <li>Presiona ⚡ Procesar archivos</li>
          <li>Explora disponibilidades y descarga resultados</li>
        </ol>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Métricas ────────────────────────────────────────────────────────────────
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
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Resumen general",
    "👤 Por docente",
    "📚 Por curso",
    "🕐 Por horario",
    "✦ Consulta con IA",
])

# ══════════════════════════════════════════════════════════════════════════
# TAB 0 — Resumen general
# ══════════════════════════════════════════════════════════════════════════
with tab0:
    st.subheader("Mapa de disponibilidad general")
    st.caption("Número de docentes disponibles por día y franja horaria.")

    matriz = {dia: {h: 0 for h in HORAS} for dia in DIAS}
    for r in exitosos:
        for dia in DIAS:
            for h_slot in HORAS:
                slot_set = hora_a_set(h_slot)
                for h_doc in r["disponibilidad"].get(dia, []):
                    if slot_set & hora_a_set(h_doc):
                        matriz[dia][h_slot] += 1
                        break

    max_val = max((matriz[d][h] for d in DIAS for h in HORAS), default=1) or 1

    def bg_color(v):
        if v == 0:
            return "#F9F9F9"
        t = v / max_val
        # blanco → naranja Ulima
        r_c = int(255 - t * (255 - 240))
        g_c = int(255 - t * (255 - 90))
        b_c = int(255 - t * (255 - 34))
        return f"rgb({r_c},{g_c},{b_c})"

    def text_color(v):
        return "#FFFFFF" if v / max_val > 0.6 else NEGRO

    header = "".join(
        f"<th style='padding:6px 10px;text-align:center;background:{NEGRO};color:#fff'>{d[:3]}</th>"
        for d in DIAS
    )
    rows = ""
    for h in HORAS:
        row_cells = ""
        for d in DIAS:
            v = matriz[d][h]
            bg = bg_color(v)
            tc = text_color(v)
            row_cells += (
                f"<td style='text-align:center;padding:5px 8px;background:{bg};"
                f"color:{tc};font-weight:600'>{v if v > 0 else '—'}</td>"
            )
        rows += (
            f"<tr><td style='padding:5px 10px;font-size:0.82rem;color:#777;"
            f"white-space:nowrap'>{h}h</td>{row_cells}</tr>"
        )

    st.markdown(f"""
    <div style="overflow-x:auto">
    <table style="border-collapse:collapse;width:100%;font-size:0.85rem">
      <thead>
        <tr>
          <th style="padding:6px 10px;background:{NEGRO};color:#fff">Franja</th>
          {header}
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
    </div>
    <p style="font-size:0.78rem;color:#999;margin-top:0.5rem">
      Naranja intenso = más docentes disponibles en esa franja.
    </p>
    """, unsafe_allow_html=True)

    st.divider()
    st.markdown("#### 🏆 Franjas con mayor disponibilidad")
    franjas_total = [
        (matriz[d][h], d, h)
        for d in DIAS for h in HORAS
        if matriz[d][h] > 0
    ]
    franjas_total.sort(reverse=True)
    if franjas_total:
        cols = st.columns(min(3, len(franjas_total)))
        for i, (v, d, h) in enumerate(franjas_total[:3]):
            with cols[i]:
                st.metric(f"{d[:3]} · {h}h", f"{v} docentes")
    else:
        st.info("Sin datos de disponibilidad registrados.")

# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — Por docente
# ══════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("¿Qué puede enseñar un docente y cuándo?")
    nombres_lista = sorted([r["nombre"] or r["archivo"] for r in exitosos])
    sel = st.selectbox("Selecciona un docente", ["— elige uno —"] + nombres_lista, key="tab1_profesor")

    if sel != "— elige uno —":
        reg = next((r for r in exitosos if (r["nombre"] or r["archivo"]) == sel), None)
        if reg:
            c1, c2 = st.columns(2)
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
            color = "card-ok" if tipo == "ASIGNADO" else "card"
            badge = "🟠 Asignado" if tipo == "ASIGNADO" else "● Disponible"
            nombre_display = (reg["nombre"] or reg["archivo"]).title()
            dias_str = "  ·  ".join(
                f"{d}: {reg['disponibilidad'][d][0]}–{reg['disponibilidad'][d][-1].split('-')[1]}"
                for d in DIAS if reg["disponibilidad"].get(d)
            ) or "Sin disponibilidad registrada"
            st.markdown(f"""
            <div class="card {color}">
              <strong>{nombre_display}</strong> &nbsp; <small>{badge}</small><br>
              <small style="color:#666">{dias_str}</small>
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
    disponibles, no_disponibles = [], []
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
                <div class="card card-ok">
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
    st.subheader("Consulta en lenguaje natural")

    api_key_activa = st.session_state.get("api_key", "")

    if not api_key_activa:
        st.info("🔑 Ingresa tu API Key en el panel izquierdo para activar el chat.")
    else:
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

        for msg in st.session_state["chat_history"]:
            if msg["role"] == "user":
                st.markdown(f'<div class="chat-user">👤 {msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="chat-bot">🤖 {msg["content"]}</div>', unsafe_allow_html=True)

        with st.form("chat_form", clear_on_submit=True):
            pregunta = st.text_input(
                "Tu pregunta",
                placeholder="Ej: ¿Quién puede dar Storytelling los martes?",
                label_visibility="collapsed",
            )
            enviado = st.form_submit_button("Enviar →", use_container_width=True)

        if enviado and pregunta.strip():
            st.session_state["chat_history"].append({"role": "user", "content": pregunta})
            with st.spinner("Consultando..."):
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key_activa)
                    messages = [{"role": m["role"], "content": m["content"]}
                                for m in st.session_state["chat_history"]]
                    respuesta = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=1024,
                        system=SYSTEM_PROMPT,
                        messages=messages,
                    )
                    texto = respuesta.content[0].text
                except ImportError:
                    texto = "⚠️ Falta instalar anthropic. Corre: `pip3 install anthropic`"
                except Exception as e:
                    texto = f"⚠️ Error: {str(e)}"
            st.session_state["chat_history"].append({"role": "assistant", "content": texto})
            st.rerun()

        if st.session_state["chat_history"]:
            if st.button("🗑 Limpiar conversación", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

        st.divider()
        st.markdown("**Ejemplos de preguntas:**")
        for ej in [
            "¿Quién puede dictar Storytelling y cuándo está disponible?",
            "¿Qué cursos puede enseñar Merino y en qué horarios?",
            "¿Quién tiene disponibilidad los lunes de 9 a 13?",
            "¿Hay alguien para Narrativa Audiovisual los martes?",
        ]:
            st.markdown(f"- *{ej}*")

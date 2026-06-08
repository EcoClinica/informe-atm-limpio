import streamlit as st
import streamlit.components.v1 as components
import datetime
import io
import re

# Intentamos importar python-docx. Si no está, avisamos amablemente.
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    st.error("Por favor, asegúrate de tener 'python-docx' en tu archivo requirements.txt")

# Configuración de pantalla ancha (Ideal para ver las 2 ATM en paralelo en el iPad)
st.set_page_config(page_title="Generador de Informes ATM", layout="wide")

# Estilos estéticos para la pantalla del iPad
st.markdown("""
    <style>
    .titulo-principal { color: #1E3A8A; font-weight: bold; text-align: center; margin-bottom: 5px; }
    .sub-titulo { text-align: center; color: #475569; font-size: 14px; margin-bottom: 25px; }
    .titulo-lado { color: #0284C7; font-size: 18px; font-weight: bold; border-bottom: 2px solid #0284C7; padding-bottom: 6px; margin-bottom: 15px; text-transform: uppercase; }
    .sub-bloque { font-weight: bold; color: #1E3A8A; margin-top: 10px; margin-bottom: 5px; font-size: 14px; }
    .resultado-calculo { background-color: #E0F2FE; padding: 10px; border-radius: 4px; border-left: 4px solid #0284C7; margin-top: 10px; font-size: 14px; color: #1E3A8A !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='titulo-principal'>Informe Ecográfico ATM</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-titulo'>Generador Profesional Automatizado con python-docx</p>", unsafe_allow_html=True)

# --- RECEPTOR DE DICTADO POR VOZ (COMPONENTE SEGURO) ---
def componente_microfono_visible(lado_id):
    js_code = f"""
    <div style="font-family: sans-serif; display: flex; flex-direction: column; gap: 4px;">
        <div style="display: flex; gap: 8px; align-items: center;">
            <button id="btn_{lado_id}" class="btn-voz btn-azul" onclick="conmutarMicro('{lado_id}')" style="flex-shrink: 0;">🎙️ Dictar 3 Medidas</button>
            <input type="text" id="output_local_{lado_id}" placeholder="Esperando dictado..." readonly 
                   style="flex-grow: 1; padding: 6px; font-size: 14px; font-weight: bold; border: 2px solid #0284C7; border-radius: 4px; background-color: #FFFFFF; color: #000000; text-align: center;">
        </div>
        <p id="status_{lado_id}" style="font-size:11px; color:#666; margin: 1px 0 0 2px; height: 14px; overflow: hidden;">Listo</p>
    </div>

    <script>
    let recognition_{lado_id} = null;
    let activo_{lado_id} = false;

    function enviarAStreamlit(textoNumeros) {{
        if (window.Streamlit) {{ Streamlit.setComponentValue(textoNumeros); }}
    }}

    function conmutarMicro(lado) {{
        const btn = document.getElementById('btn_' + lado);
        const status = document.getElementById('status_' + lado);
        const inputLocal = document.getElementById('output_local_' + lado);

        if (activo_{lado_id}) {{
            if (recognition_{lado_id}) recognition_{lado_id}.abort();
            resetearBoton(lado, "🛑 Cancelado.");
            return;
        }}

        if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {{
            status.innerText = "No compatible.";
            return;
        }}

        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        recognition_{lado_id} = new SpeechRecognition();
        recognition_{lado_id}.lang = 'es-ES';
        recognition_{lado_id}.interimResults = false;

        activo_{lado_id} = true;
        btn.innerText = "🛑 Parar";
        btn.className = "btn-voz btn-rojo";
        status.innerText = "🎙️ Escuchando...";

        recognition_{lado_id}.start();

        recognition_{lado_id}.onresult = function(event) {{
            const texto = event.results[0][0].transcript;
            const matches = texto.replace(/,/g, '.').match(/[0-9]+(\\.[0-9]+)?/g);
            
            if (matches && matches.length >= 3) {{
                const resultadoCadena = matches[0] + " , " + matches[1] + " , " + matches[2];
                inputLocal.value = resultadoCadena;
                status.innerText = "✓ Capturadas.";
                status.style.color = "#16A34A";
                enviarAStreamlit(resultadoCadena);
            }} else {{
                status.innerText = "❌ Reintenta.";
                status.style.color = "#DC2626";
            }}
        }};

        recognition_{lado_id}.onerror = function() {{ resetearBoton(lado, "Reintenta."); }};
        recognition_{lado_id}.onend = function() {{ resetearBoton(lado, status.innerText); }};
    }}

    function resetearBoton(lado, msg) {{
        activo_{lado_id} = false;
        const btn = document.getElementById('btn_' + lado);
        const status = document.getElementById('status_' + lado);
        btn.innerText = "🎙️ Dictar 3 Medidas";
        btn.className = "btn-voz btn-azul";
        if(msg) status.innerText = msg;
    }}
    </script>
    <style>
    .btn-voz {{ border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px; height: 32px; }}
    .btn-azul {{ background-color: #0284C7; color: white; }}
    .btn-rojo {{ background-color: #DC2626; color: white; }}
    </style>
    """
    return components.html(js_code, height=60, scrolling=False)

# --- PROCESADORES NUMÉRICOS ---
def procesar_medidas_sistema(texto_dictado, manual_as, manual_lat, manual_pi):
    if texto_dictado and isinstance(texto_dictado, str) and len(texto_dictado.strip()) > 0:
        numeros = re.findall(r"[0-9]+(?:\.[0-9]+)?", texto_dictado)
        if len(numeros) >= 3:
            return numeros[0], numeros[1], numeros[2]
    return manual_as, manual_lat, manual_pi

def calcular_posicion_condilo(ant_sup_txt, post_inf_txt):
    try:
        if not ant_sup_txt or not post_inf_txt: return "0.00"
        as_val = float(str(ant_sup_txt).replace(',', '.'))
        pi_val = float(str(post_inf_txt).replace(',', '.'))
        if (pi_val + as_val) == 0: return "0.00"
        resultado = ((pi_val - as_val) / (pi_val + as_val)) * 100
        return f"{'+' if resultado > 0 else ''}{resultado:.2f}"
    except ValueError:
        return "0.00"

# --- RECOLECCIÓN DE DATOS EN LA INTERFAZ WEB ---
st.subheader("📋 Datos del Paciente")
cp1, cp2, cp3 = st.columns(3)
with cp1:
    nombres = st.text_input("Nombre del Paciente:", value="")
    apellidos = st.text_input("Apellidos:", value="")
with cp2:
    edad = st.text_input("Edad:")
    fecha = st.date_input("Fecha del estudio:", datetime.date.today(), format="DD/MM/YYYY")
with cp3:
    motivo = st.text_input("Motivo de consulta:", value="Dolor ATM bilateral.")
    derivado = st.text_input("Derivado por (Dr/a):", value="")

# Opciones de los menús desplegables rápidos
opts_morfologia = ["aplanado, irregular", "irregular, estrecho, con cresta lateral", "redondeado, regular", "en pico de pájaro"]
opts_espacio = ["con osteofitos. Engrosamiento sinovial superior", "con osteofitos", "libre, sin hallazgos patológicos"]
opts_derrame = ["Presencia de derrame articular.", "Sin presencia de derrame articular."]
opts_ecoestructura = ["Ecoestructura hipoecogénico.", "Ecoestructura homogénea.", "Ecoestructura heterogénea."]
opts_horas = ["hora 11", "hora 12", "hora 10", "hora 1"]
opts_relacion = ["Cóndilo en posición anterior", "Cóndilo en posición central", "Cóndilo en posición posterior"]
opts_cerrada = ["en hora 11 cubre parcialmente la cabeza del cóndilo.", "cubre totalmente la cabeza del cóndilo.", "desplazamiento, no cubre la cabeza condilar."]
opts_abierta = ["desplazamiento anterior cubre la porción anterior de la cabeza condilar durante la apertura bucal, resto del cóndilo contacta con la cavidad glenoidea.", "desplazamiento discal normal, cubre por completo la cabeza del cóndilo."]

st.markdown("<br>", unsafe_allow_html=True)

# --- INTERFAZ VISUAL EN PARALELO (DERECHA / IZQUIERDA) ---
col_der, col_izq = st.columns(2, gap="large")

# LADO DERECHO
with col_der:
    st.markdown("<div class='titulo-lado'>Estudio ATM Derecha</div>", unsafe_allow_html=True)
    morf_der = st.selectbox("Morfología cabeza condilar (D):", opts_morfologia, index=0, key="w_m_der")
    esp_der = st.selectbox("Espacio articular (D):", opts_espacio, index=0, key="w_e_der")
    derrame_der = st.selectbox("Derrame articular (D):", opts_derrame, index=0, key="w_d_der")
    
    st.markdown("<div class='sub-bloque'>Medidas Cóndilo Derecho:</div>", unsafe_allow_html=True)
    dictado_der = componente_microfono_visible("der")
    md1, md2, md3 = st.columns(3)
    with md1: manual_as_der = st.text_input("Ant-Sup (D)", value="1.98", key="w_as_der")
    with md2: manual_lat_der = st.text_input("Lateral (D)", value="1.65", key="w_la_der")
    with md3: manual_pi_der = st.text_input("Post-Inf (D)", value="2.84", key="w_pi_der")
    
    med_as_der, med_lat_der, med_pi_der = procesar_medidas_sistema(dictado_der, manual_as_der, manual_lat_der, manual_pi_der)
    res_der = calcular_posicion_condilo(med_as_der, med_pi_der)
    st.markdown(f"<div class='resultado-calculo'>🧮 Índice de Pullinger (D): {res_der}%</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='sub-bloque'>Disco Articular y Dinámica (D):</div>", unsafe_allow_html=True)
    eco_der = st.selectbox("Ecoestructura (D):", opts_ecoestructura, index=0, key="w_ec_der")
    hora_der = st.selectbox("Situación (D):", opts_horas, index=0, key="w_h_der")
    rel_der = st.selectbox("Relación cóndilo-fosa (D):", opts_relacion, index=0, key="w_r_der")
    c_der = st.selectbox("Boca cerrada (D):", opts_cerrada, index=0, key="w_c_der")
    a_der = st.selectbox("Boca abierta (D):", opts_abierta, index=0, key="w_a_der")

# LADO IZQUIERDO
with col_izq:
    st.markdown("<div class='titulo-lado'>Estudio ATM Izquierda</div>", unsafe_allow_html=True)
    morf_izq = st.selectbox("Morfología cabeza condilar (I):", opts_morfologia, index=1, key="w_m_izq")
    esp_izq = st.selectbox("Espacio articular (I):", opts_espacio, index=1, key="w_e_izq")
    derrame_izq = st.selectbox("Derrame articular (I):", opts_derrame, index=0, key="w_d_izq")
    
    st.markdown("<div class='sub-bloque'>Medidas Cóndilo Izquierdo:</div>", unsafe_allow_html=True)
    dictado_izq = componente_microfono_visible("izq")
    mi1, mi2, mi3 = st.columns(3)
    with mi1: manual_as_izq = st.text_input("Ant-Sup (I)", value="2.37", key="w_as_izq")
    with mi2: manual_lat_izq = st.text_input("Lateral (I)", value="1.14", key="w_la_izq")
    with mi3: manual_pi_izq = st.text_input("Post-Inf (I)", value="2.92", key="w_pi_izq")
    
    med_as_izq, med_lat_izq, med_pi_izq = procesar_medidas_sistema(dictado_izq, manual_as_izq, manual_lat_izq, manual_pi_izq)
    res_izq = calcular_posicion_condilo(med_as_izq, med_pi_izq)
    st.markdown(f"<div class='resultado-calculo'>🧮 Índice de Pullinger (I): {res_izq}%</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='sub-bloque'>Disco Articular y Dinámica (I):</div>", unsafe_allow_html=True)
    eco_izq = st.selectbox("Ecoestructura (I):", opts_ecoestructura, index=0, key="w_ec_izq")
    hora_izq = st.selectbox("Situación (I):", opts_horas, index=0, key="w_h_izq")
    rel_izq = st.selectbox("Relación cóndilo-fosa (I):", opts_relacion, index=1, key="w_r_izq")
    c_izq = st.selectbox("Boca cerrada (I):", opts_cerrada, index=0, key="w_c_izq")
    a_izq = st.selectbox("Boca abierta (I):", opts_abierta, index=0, key="w_a_izq")

st.markdown("<br><hr>", unsafe_allow_html=True)
conclusion = st.text_area("📝 CONCLUSIÓN DEL INFORME:", value="Signos ecográficos descritos compatibles con patología articular bilateral.")

# --- PASO 1 Y 2 UNIFICADOS: EL MOTOR QUE DISEÑA EL WORD INTERNAMENTE ---
st.subheader("💾 Compilar y Descargar Informe Médico")

if st.button("🚀 COMPILAR INFORME EN FORMATO PROFESIONAL (.DOCX)"):
    try:
        # Creamos el documento en blanco
        doc = Document()
        
        # Configuración de márgenes
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # Configuración de Estilo Base (Fuente Arial, Tamaño 11)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        font.color.rgb = RGBColor(0x33, 0x41, 0x55) # Gris oscuro elegante

        # --- TITULO PRINCIPAL DEL WORD ---
        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_tit = p_tit.add_run("INFORME ECOGRÁFICO DE LA ARTICULACIÓN TEMPOROMANDIBULAR (ATM)")
        run_tit.font.size = Pt(14)
        run_tit.font.bold = True
        run_tit.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A) # Azul Marino Médico

        # Subtítulo fijo institucional
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = p_sub.add_run("Protocolo de adquisición ecográfica en posición de boca abierta y cerrada (ambas)")
        run_sub.font.size = Pt(10)
        run_sub.font.italic = True
        
        doc.add_paragraph("-" * 80) # Separador visual

        # --- BLOQUE: DATOS GENERALES DEL PACIENTE ---
        p_dat = doc.add_paragraph()
        p_dat.add_run("Paciente: ").bold = True
        p_dat.add_run(f"{nombres} {apellidos}         ")
        p_dat.add_run("Edad: ").bold = True
        p_dat.add_run(f"{edad} años\n")
        p_dat.add_run("Fecha: ").bold = True
        p_dat.add_run(f"{fecha.strftime('%d/%m/%Y')}         ")
        p_dat.add_run("Derivado por: ").bold = True
        p_dat.add_run(f"{derivado}\n")
        p_dat.add_run("Motivo de consulta: ").bold = True
        p_dat.add_run(f"{motivo}")

        doc.add_paragraph("-" * 80)

        # --- SECCIÓN: ARTICULACIÓN DERECHA ---
        h_der = doc.add_paragraph()
        r_h_der = h_der.add_run("ESTUDIO ARTICULACIÓN TEMPOROMANDIBULAR DERECHA")
        r_h_der.font.bold = True
        r_h_der.font.size = Pt(12)
        r_h_der.font.color.rgb = RGBColor(0x02, 0x84, 0xC7) # Azul Celeste Clínico

        p_body_der = doc.add_paragraph()
        p_body_der.add_run("Cóndilo mandibular derecho: ").bold = True
        p_body_der.add_run(f"morfología cabeza condilar {morf_der}.\n")
        p_body_der.add_run("Espacio articular: ").bold = True
        p_body_der.add_run(f"{esp_der}.\n")
        
        # Línea de medidas destacada
        p_body_der.add_run(f"Medidas anterosuperior: {med_as_der} mm. Lateral: {med_lat_der} mm. Posteroinferior: {med_pi_der} mm.\n").bold = True
        p_body_der.add_run(f"{derrame_der} (Índice Condilar de Pullinger: {res_der}%)\n")
        
        p_body_der.add_run("Disco articular: ").bold = True
        p_body_der.add_run(f"{eco_der} Situación en {hora_der}.\n")
        p_body_der.add_run("Relación cóndilo-fosa: ").bold = True
        p_body_der.add_run(f"{rel_der}.\n\n")
        
        p_body_der.add_run("Estudio dinámico (Posición del disco derecho):\n").bold = True
        doc.add_paragraph(f"Boca cerrada: en {hora_der} {c_der}", style='List Bullet')
        doc.add_paragraph(f"Boca abierta: {a_der}", style='List Bullet')

        doc.add_paragraph(" ") # Espacio de separación

        # --- SECCIÓN: ARTICULACIÓN IZQUIERDA ---
        h_izq = doc.add_paragraph()
        r_h_izq = h_izq.add_run("ESTUDIO ARTICULACIÓN TEMPOROMANDIBULAR IZQUIERDA")
        r_h_izq.font.bold = True
        r_h_izq.font.size = Pt(12)
        r_h_izq.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

        p_body_izq = doc.add_paragraph()
        p_body_izq.add_run("Cóndilo mandibular: ").bold = True
        p_body_izq.add_run(f"la morfología de la cabeza condilar {morf_izq}.\n")
        p_body_izq.add_run("Espacio articular: ").bold = True
        p_body_izq.add_run(f"{esp_izq}.\n")
        
        # Línea de medidas destacada
        p_body_izq.add_run(f"Medidas Anterosuperior: {med_as_izq} mm. Lateral: {med_lat_izq} mm. Posteroinferior: {med_pi_izq} mm.\n").bold = True
        p_body_izq.add_run(f"{derrame_izq} (Índice Condilar de Pullinger: {res_izq}%)\n")
        
        p_body_izq.add_run("Disco articular: ").bold = True
        p_body_izq.add_run(f"{eco_izq} Situado en {hora_izq}.\n")
        p_body_izq.add_run("Relación cóndilo-fosa: ").bold = True
        p_body_izq.add_run(f"{rel_izq}.\n\n")
        
        p_body_izq.add_run("Estudio dinámico (Posición del disco izquierdo):\n").bold = True
        doc.add_paragraph(f"Boca cerrada: en {hora_izq} {c_izq}", style='List Bullet')
        doc.add_paragraph(f"Boca abierta: {a_izq}", style='List Bullet')

        doc.add_paragraph("-" * 80)

        # --- SECCIÓN: CONCLUSIÓN FINAL ---
        p_con = doc.add_paragraph()
        r_con_title = p_con.add_run("CONCLUSIÓN: ")
        r_con_title.bold = True
        r_con_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        p_con.add_run(conclusion)

        # Guardado en memoria intermedia para descarga inmediata
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        nombre_limpio = f"Informe_ATM_{apellidos if apellidos else 'Paciente'}.docx".replace(" ", "_")
        
        st.download_button(
            label="⬇️ DESCARGAR DOCUMENTO WORD FORMATEADO",
            data=bio,
            file_name=nombre_limpio,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success("¡Documento Word generado con éxito desde el código! Listo para descargar.")
        
    except Exception as e:
        st.error(f"Error crítico de formateo en el motor de python-docx: {e}")

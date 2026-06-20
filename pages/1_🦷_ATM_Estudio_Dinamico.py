import streamlit as st
import streamlit.components.v1 as components
from docxtpl import DocxTemplate
import datetime
import io
import re
import os
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generar_plantilla_fiel():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    AZUL_CLARO = RGBColor(2, 132, 199)
    NEGRO = RGBColor(0, 0, 0)
    GRIS_LINEA = RGBColor(156, 163, 175)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("INFORME ECOGRÁFICO DE LA ARTICULACIÓN TEMPOROMANDIBULAR\n(ATM)")
    r.bold = True
    r.font.name = 'Arial'
    r.font.size = Pt(14)
    r.font.color.rgb = NEGRO
    
    p_linea = doc.add_paragraph()
    r_l1 = p_linea.add_run("--------------------------------------------------------------------------------")
    r_l1.font.name = 'Arial'
    r_l1.font.color.rgb = GRIS_LINEA
    
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    
    r = p.add_run("Paciente: ")
    r.bold = True
    r.font.name = 'Arial'
    p.add_run("{{ paciente }}                                                                             ").font.name = 'Arial'
    
    p.add_run("Edad: ").bold = True
    p.add_run("{{ edad }}\n").font.name = 'Arial'
    
    p.add_run("Fecha: ").bold = True
    p.add_run("{{ fecha }}                                                                     ").font.name = 'Arial'
    
    p.add_run("Derivado por: ").bold = True
    p.add_run("{{ derivado }}\n").font.name = 'Arial'
    
    p.add_run("Motivo de consulta: ").bold = True
    p.add_run("{{ motivo }}").font.name = 'Arial'
    
    p_linea_desc = doc.add_paragraph()
    p_linea_desc.paragraph_format.space_before = Pt(6)
    p_linea_desc.paragraph_format.space_after = Pt(12)
    r_ldesc = p_linea_desc.add_run("--------------------------------------------------------------------------------")
    r_ldesc.font.name = 'Arial'
    r_ldesc.font.color.rgb = GRIS_LINEA
    
    p_estudio = doc.add_paragraph()
    p_estudio.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_estudio.add_run("ESTUDIO DINÁMICO AMBAS ATM")
    r.bold = True
    r.font.name = 'Arial'
    r.font.size = Pt(12)
    
    def agregar_bloque_atm(lado_nombre, prefijo):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        r = p.add_run(f"ESTUDIO ARTICULACIÓN TEMPOROMANDIBULAR {lado_nombre}")
        r.bold = True
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        r.font.color.rgb = AZUL_CLARO
        
        p_campos = doc.add_paragraph()
        campos = [("Cóndilo mandibular: ", f"{{{{ condilo_{prefijo} }}}}\n"), 
                  ("Espacio articular: ", f"{{{{ espacio_{prefijo} }}}}\n"), 
                  ("Derrame articular: ", f"{{{{ derrame_{prefijo} }}}}\n")]
        for l, v in campos:
            p_campos.add_run(l).bold = True
            p_campos.add_run(v).font.name = 'Arial'
            
        p_med = doc.add_paragraph()
        p_med.add_run("Medidas condilares: ").bold = True
        p_med.add_run("Anterior ").italic = True
        p_med.add_run(f"{{{{ med_as_{prefijo} }}}} mm.  ")
        p_med.add_run("Lateral: ").italic = True
        p_med.add_run(f"{{{{ med_lat_{prefijo} }}}} mm.  ")
        p_med.add_run("Posterior: ").italic = True
        p_med.add_run(f"{{{{ med_pi_{prefijo} }}}} mm.\n")
        
        p_med.add_run("Posición condilar (Pullinger): ").bold = True
        p_med.add_run(f"{{{{ pullinger_{prefijo} }}}}\n")
        p_med.add_run("Relación cóndilo-fosa glenoidea: ").bold = True
        p_med.add_run(f"{{{{ relacion_{prefijo} }}}}")
        
        p_disc = doc.add_paragraph()
        p_disc.add_run("Disco articular: ").bold = True
        p_disc.add_run(f"{{{{ disco_{prefijo} }}}}                                                           ")
        p_disc.add_run("Situación en hora: ").bold = True
        p_disc.add_run(f"{{{{ hora_{prefijo} }}}}")
        
        p_din = doc.add_paragraph()
        p_din.add_run("Estudio dinámico:\n").bold = True
        p_din.add_run("·       Boca cerrada: ").bold = True
        p_din.add_run(f"{{{{ cerrada_{prefijo} }}}}\n")
        p_din.add_run("·       Boca abierta: ").bold = True
        p_din.add_run(f"{{{{ abierta_{prefijo} }}}}\n")
        p_din.add_run("·       Reposición: ").bold = True
        p_din.add_run(f"{{{{ repo_{prefijo} }}}}")
        
    agregar_bloque_atm("DERECHA", "der")
    agregar_bloque_atm("IZQUIERDA", "izq")
    
    p_linea2 = doc.add_paragraph()
    p_linea2.add_run("--------------------------------------------------------------------------------").font.color.rgb = GRIS_LINEA
    
    p_c = doc.add_paragraph()
    r = p_c.add_run("CONCLUSIÓN: ")
    r.bold = True
    r.font.color.rgb = AZUL_CLARO
    p_c.add_run("{{ conclusion }}")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

try:
    with open("plantilla_atm.docx", "wb") as f: f.write(generar_plantilla_fiel().getbuffer())
except: pass

st.set_page_config(page_title="Informe Ecográfico ATM", layout="wide", page_icon="🎙️")

st.markdown("""
    <style>
    .titulo-principal { color: #1E3A8A; font-weight: bold; text-align: center; margin-top: -20px; }
    .sub-seccion { color: #0284C7; border-bottom: 2px solid #0284C7; padding-bottom: 5px; margin-bottom: 15px; font-size: 20px; }
    .titulo-medidas { font-size: 14px; font-weight: bold; margin-bottom: 5px; color: #1E3A8A !important; }
    .resultado-calculo { background-color: #E0F2FE; padding: 12px; border-radius: 5px; border-left: 4px solid #0284C7; margin-top: 10px; margin-bottom: 15px; font-size: 14px; color: #1E3A8A !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.sidebar.markdown("### 📂 Zona de Respaldos")
st.sidebar.download_button("📥 Descargar plantilla original (.docx)", generar_plantilla_fiel(), "plantilla_atm.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("<h1 class='titulo-principal'>Informe Ecográfico de la Articulación Temporomandibular (ATM)</h1>", unsafe_allow_html=True)

if "dictado_der" not in st.session_state: st.session_state.dictado_der = ""
if "dictado_izq" not in st.session_state: st.session_state.dictado_izq = ""

def componente_microfono_visible(lado_id):
    js_code = f"""
    <div style="font-family: sans-serif; display: flex; gap: 10px; align-items: center;">
        <button id="btn_{lado_id}" onclick="conmutarMicro('{lado_id}')" style="background-color:#0284C7; color:white; border:none; padding:6px 12px; border-radius:4px; font-weight:bold; cursor:pointer;">🎙️ Dictar 3 Medidas</button>
        <input type="text" id="out_{lado_id}" readonly style="padding:4px; border:2px solid #0284C7; border-radius:4px; text-align:center; width:150px;">
    </div>
    <script>
    let rec_{lado_id} = null; let act_{lado_id} = false;
    function conmutarMicro(l) {{
        const btn = document.getElementById('btn_'+l); const inp = document.getElementById('out_'+l);
        if(act_{lado_id}) {{ rec_{lado_id}.abort(); return; }}
        const SR = window.SpeechRecognition || window.webkitSpeechRecognition; if(!SR) return;
        rec_{lado_id} = new SR(); rec_{lado_id}.lang='es-ES'; act_{lado_id}=true; btn.innerText="🛑 Parar";
        rec_{lado_id}.start();
        rec_{lado_id}.onresult = function(e) {{
            const t = e.results[0][0].transcript.replace(/,/g, '.');
            const m = t.match(/[0-9]+(\\.[0-9]+)?/g);
            if(m && m.length >= 3) {{ inp.value = m[0]+" , "+m[1]+" , "+m[2]; if(window.Streamlit) Streamlit.setComponentValue(inp.value); }}
        }};
        rec_{lado_id}.onend = function() {{ act_{lado_id}=false; btn.innerText="🎙️ Dictar 3 Medidas"; }};
    }}
    (function() {{
        var s = document.createElement('script'); s.src = "https://cdn.jsdelivr.net/npm/@streamlit/component-lib@1.4.0/dist/index.min.js";
        s.onload = function() {{ window.addEventListener('load', function() {{ Streamlit.setFrameHeight(40); }}); }}; document.head.appendChild(s);
    }})();
    </script>
    """
    return components.html(js_code, height=40)

st.subheader("📋 Datos del Paciente")
with st.container(border=True):
    cp1, cp2, cp3 = st.columns(3)
    with cp1:
        nombres = st.text_input("Paciente:")
        edad = st.text_input("Edad:")
    with cp2:
        fecha = st.date_input("Fecha:", datetime.date.today(), format="DD/MM/YYYY")
        derivado = st.text_input("Derivado por:")
    with cp3:
        motivo = st.text_input("Motivo de consulta:")

opts_condilo = ["", "Redondeado", "Aplanado", "En pico de pájaro (en punta)", "Con cresta central", "Con cresta marginal"]
opts_espacio = ["", "Libre", "Con engrosamiento sinovial", "Osteofitos", "Regular", "Irregular"]
opts_derrame = ["", "Sin derrame articular", "Con derrame anecoico", "Con derrame articular y con partículas ecogénicas"]
opts_relacion = ["", "Central", "Anterior", "Posterior"]
opts_disco = ["", "Homogénea, hipoecogénico", "Heterogénea", "Irregular"]
opts_horas = ["", "12", "11", "10", "1", "2"]
opts_boca_cerrada = ["", "Cubre totalmente la cabeza del cóndilo", "Cubre parcialmente la cabeza del cóndilo", "Desplazamiento, no cubre la cabeza condilar"]
opts_boca_abierta = ["", "Desplazamiento discal normal, cubre la cabeza del cóndilo", "Desplazamiento anterior, el disco cubre parcialmente la cabeza del cóndilo", "Desplazamiento anterior con subluxación discal", "Desplazamiento anterior con luxación discal"]
opts_repo = ["", "Espontánea", "Requiere maniobras del paciente", "Requiere maniobras del médico", "No se reposiciona"]

col_der, col_izq = st.columns(2)

with col_der:
    with st.container(border=True):
        st.markdown("<h2 class='sub-seccion'>🔹ATM Derecha</h2>", unsafe_allow_html=True)
        condilo_der = st.selectbox("Cóndilo (D):", opts_condilo, key="c_d")
        espacio_der = st.selectbox("Espacio (D):", opts_espacio, key="e_d")
        derrame_der = st.selectbox("Derrame (D):", opts_derrame, key="d_d")
        
        st.markdown("<p class='titulo-medidas'>Medidas condilares (mm):</p>", unsafe_allow_html=True)
        r_md = componente_microfono_visible("der")
        if r_md: st.session_state.dictado_der = r_md
        
        m1, m2, m3 = st.columns(3)
        with m1: ma_d = st.text_input("Anterior (D)", key="ma_d")
        with m2: ml_d = st.text_input("Lateral (D)", key="ml_d")
        with m3: mp_d = st.text_input("Posterior (D)", key="mp_d")
        
       def proc(d, a, l, p):
            # Convertimos "d" a texto de forma segura para evitar el TypeError
            texto_dictado = str(d) if d is not None else ""
            if texto_dictado and len(re.findall(r"[0-9]+", texto_dictado)) >= 3: 
                return re.findall(r"[0-9.]+", texto_dictado)[:3]
            return a, l, p
        
        def calc(a, p):
            try:
                if not a or not p: return "Pendiente"
                v_a, v_p = float(str(a).replace(',','.')), float(str(p).replace(',','.'))
                return f"{'+' if (v_p-v_a)>0 else ''}{((v_p-v_a)/(v_p+v_a))*100:.2f}%"
            except: return "Error"
        res_d = calc(v1, v3)
        st.markdown(f"<div class='resultado-calculo'>🧮 Pullinger (D): {res_d}</div>", unsafe_allow_html=True)
        
        rel_d = st.selectbox("Relación (D):", opts_relacion, key="r_d")
        disc_d = st.selectbox("Disco (D):", opts_disco, key="di_d")
        h_d = st.selectbox("Hora (D):", opts_horas, key="h_d")
        cerr_d = st.selectbox("Boca cerrada (D):", opts_boca_cerrada, key="ce_d")
        ab_d = st.selectbox("Boca abierta (D):", opts_boca_abierta, key="ab_d")
        rep_d = st.selectbox("Reposición (D):", opts_repo, key="re_d")

with col_izq:
    with st.container(border=True):
        st.markdown("<h2 class='sub-seccion'>🔹ATM Izquierda</h2>", unsafe_allow_html=True)
        condilo_izq = st.selectbox("Cóndilo (I):", opts_condilo, key="c_i")
        espacio_izq = st.selectbox("Espacio (I):", opts_espacio, key="e_i")
        derrame_izq = st.selectbox("Derrame (I):", opts_derrame, key="d_i")
        
        st.markdown("<p class='titulo-medidas'>Medidas condilares (mm):</p>", unsafe_allow_html=True)
        r_mi = componente_microfono_visible("izq")
        if r_mi: st.session_state.dictado_izq = r_mi
        
        m4, m5, m6 = st.columns(3)
        with m4: ma_i = st.text_input("Anterior (I)", key="ma_i")
        with m5: ml_i = st.text_input("Lateral (I)", key="ml_i")
        with m6: mp_i = st.text_input("Posterior (I)", key="mp_i")
        v4, v5, v6 = proc(st.session_state.dictado_izq, ma_i, ml_i, mp_i)
        res_i = calc(v4, v6)
        st.markdown(f"<div class='resultado-calculo'>🧮 Pullinger (I): {res_i}</div>", unsafe_allow_html=True)
        
        rel_i = st.selectbox("Relación (I):", opts_relacion, key="r_i")
        disc_i = st.selectbox("Disco (I):", opts_disco, key="di_i")
        h_i = st.selectbox("Hora (I):", opts_horas, key="h_i")
        cerr_i = st.selectbox("Boca cerrada (I):", opts_boca_cerrada, key="ce_i")
        ab_i = st.selectbox("Boca abierta (I):", opts_boca_abierta, key="ab_i")
        rep_i = st.selectbox("Reposición (I):", opts_repo, key="re_i")

st.markdown("<br>", unsafe_allow_html=True)
conclusion = st.text_area("📝 CONCLUSIÓN:")

try:
    doc = DocxTemplate("plantilla_atm.docx")
    ctx = {
        'paciente': nombres, 'edad': edad, 'derivado': derivado, 'fecha': fecha.strftime("%d/%m/%Y"), 'motivo': motivo,
        'condilo_der': condilo_der, 'espacio_der': espacio_der, 'derrame_der': derrame_der, 'med_as_der': v1, 'med_lat_der': v2, 'med_pi_der': v3, 'pullinger_der': res_d, 'relacion_der': rel_d, 'disco_der': disc_d, 'hora_der': h_d, 'cerrada_der': cerr_d, 'abierta_der': ab_d, 'repo_der': rep_d,
        'condilo_izq': condilo_izq, 'espacio_izq': espacio_izq, 'derrame_izq': derrame_izq, 'med_as_izq': v4, 'med_lat_izq': v5, 'med_pi_izq': v6, 'pullinger_izq': res_i, 'relacion_izq': rel_i, 'disco_izq': disc_i, 'hora_izq': h_i, 'cerrada_izq': cerr_i, 'abierta_izq': ab_i, 'repo_izq': rep_i,
        'conclusion': conclusion
    }
    doc.render(ctx)
    bio_o = io.BytesIO()
    doc.save(bio_o)
    bio_o.seek(0)
    st.download_button("🚀 DESCARGAR INFORME EN WORD", bio_o, f"Informe_ATM_{nombres}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
except Exception as e:
    st.error(f"Error: {e}")

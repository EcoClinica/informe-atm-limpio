import streamlit as st

st.set_page_config(page_title="EcoClínica - Panel de Informes", layout="wide", page_icon="🏥")

st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>🏥 EcoClínica - Sistema de Informes</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border: 1px solid #1E3A8A;'>", unsafe_allow_html=True)

st.write("### 🩺 Panel de Selección")
st.info("👈 Por favor, despliega el menú lateral izquierdo y selecciona el grupo ecográfico que vas a realizar.")

st.markdown("""
### 📂 Grupos Disponibles de Exploración:
* **🦷 Grupo ATM y Maxilofacial**:
  - Informe 1: ATM (Estudio Dinámico)
  - Informe 2: Músculos Masticadores
* **💪 Grupo Musculoesquelético General *(Próximamente)*🏼**:
  - Informes de Cuello (ECM), Extremidades (Cuádriceps, Gemelos), etc.
""")

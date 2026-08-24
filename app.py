import streamlit as st
import datetime
import requests
import json

st.set_page_config(page_title="Reporte de Daños - Mantenimiento", page_icon="⚙️", layout="centered")

# PEGA AQUÍ ENTRE LAS COMILLAS EL ENLACE QUE COPIASTE DE GOOGLE APPS SCRIPT (el que termina en /exec)
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzKYQ_9Sx-gRo778XZLF2V_4iUaIs0IHKWTbGHc1q3dEmOI32-gaLLURA4aqE0GSjTg/exec"

# Listas predeterminadas (puedes ajustarlas según tus máquinas y técnicos reales)
maquinas = [
    "WNT", "SELCO2", "SELCO3", "SELCO4", "HOMAG400", "HOMAG500", "HOMAGKL310", 
    "STREAM1", "STREAM2", "STREAM3", "AKRON1", "AKRON2", "JADE", "NANXING", 
    "SKIPPER1", "SKIPPER2", "BHX1", "BHX2", "ROVER1", "NESTING", "VITAP"
]

tecnicos = [
    "WILLIAN DIAZ", "JAIRO ISAZA", "DAIRON MENESES", "LEONARDO ALVAREZ", 
    "BAYRON LOPEZ", "HEBERT CHAPUEL", "ESTEBAN ROSERO", "BRANDON RAMOS", 
    "DAVID PANTOJA", "CARLOS LUGO", "JULIO DAZA", "JHOAN MOTATO"
]

def generar_horas_am_pm():
    horas = []
    for h in range(1, 13):
        for m in range(0, 60, 5):
            horas.append(f"{h:02d}:{m:02d} AM")
    for h in range(1, 13):
        for m in range(0, 60, 5):
            horas.append(f"{h:02d}:{m:02d} PM")
    return horas

lista_horas = generar_horas_am_pm()

st.title("📱 Reporte Diario de Daños")
st.markdown("Registra las fallas de planta directamente hacia Google Sheets.")

with st.form("form_reporte_daño", clear_on_submit=True):
    fecha = st.date_input("Fecha del reporte", datetime.date.today())
    
    maquina = st.selectbox("Máquina / Equipo", maquinas)
    
    st.markdown("🕒 **Selección de Tiempos (AM / PM)**")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        str_hora_inicio = st.selectbox("Hora de Inicio del Paro", lista_horas, index=120)
    with col_h2:
        str_hora_fin = st.selectbox("Hora de Finalización", lista_horas, index=126)
    
    daño = st.text_area("Descripción del Daño / Falla")
    reparacion = st.text_area("Acción de Reparación Realizada")
    
    st.markdown("---")
    st.subheader("Personal de Mantenimiento y Repuestos")
    
    col3, col4, col5 = st.columns(3)
    with col3:
        tecnico1 = st.selectbox("Técnico 1", [""] + tecnicos)
    with col4:
        tecnico2 = st.selectbox("Técnico 2", [""] + tecnicos)
    with col5:
        tecnico3 = st.selectbox("Técnico 3", [""] + tecnicos)
        
    repuesto = st.text_input("Repuesto utilizado (Opcional)")
    cantidad = st.number_input("Cantidad", min_value=0, step=1)
    
    enviar = st.form_submit_button("💾 Guardar Registro en Google Sheets")
    
    if enviar:
        if not GOOGLE_SCRIPT_URL or GOOGLE_SCRIPT_URL == "PEGA_AQUI_TU_URL_DE_GOOGLE_APPS_SCRIPT":
            st.error("Por favor configura la URL de Google Apps Script en el código.")
        elif not daño.strip() or not reparacion.strip():
            st.warning("Por favor completa la descripción del daño y la reparación.")
        else:
            def parse_am_pm(t_str, fecha_base):
                t_partes, periodo = t_str.split(" ")
                h_str, m_str = t_partes.split(":")
                h = int(h_str)
                m = int(m_str)
                if periodo == "PM" and h != 12:
                    h += 12
                if periodo == "AM" and h == 12:
                    h = 0
                return datetime.datetime.combine(fecha_base, datetime.time(h, m))

            dt_ini = parse_am_pm(str_hora_inicio, fecha)
            dt_fin = parse_am_pm(str_hora_fin, fecha)
            
            if dt_fin < dt_ini:
                dt_fin += datetime.timedelta(days=1)
            
            diferencia = dt_fin - dt_ini
            horas = diferencia.seconds // 3600
            minutos = (diferencia.seconds % 3600) // 60
            tiempo_real = f"{horas:02d}:{minutos:02d}:00"
            
            lista_repara = [t for t in [tecnico1, tecnico2, tecnico3] if t != ""]
            quien_repara = ", ".join(lista_repara) if lista_repara else ""
            
            payload = {
                "fecha": str(fecha),
                "maquina": maquina,
                "hora_inicio": dt_ini.strftime("%H:%M:%S"),
                "hora_fin": dt_fin.strftime("%H:%M:%S"),
                "tiempo_real": tiempo_real,
                "daño": daño,
                "reparacion": reparacion,
                "tecnico1": tecnico1,
                "tecnico2": tecnico2,
                "tecnico3": tecnico3,
                "quien_repara": quien_repara,
                "repuesto": repuesto,
                "cantidad": str(cantidad) if repuesto else ""
            }
            
            try:
                response = requests.post(GOOGLE_SCRIPT_URL, json=payload)
                if response.status_code == 200:
                    st.success("¡Daño registrado y guardado con éxito en Google Sheets!")
                else:
                    st.error("Error al conectar con Google Sheets.")
            except Exception as e:
                st.error(f"Error de conexión: {e}")

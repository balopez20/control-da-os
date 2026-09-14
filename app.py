import streamlit as st
import pandas as pd
import datetime
import requests
import json
import os

st.set_page_config(page_title="Reporte de Daños - Mantenimiento", page_icon="⚙️", layout="centered")

GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw-ev0tuia3kl6Wkdec_0Q3-CxOlRHJ2cgDuGJp9Rf94ehBDEu4X7putnQhUsu33CGmhw/exec"

maquinas = [
    "WNT", "SELCO2", "SELCO3", "SELCO4", "HOMAG400", "HOMAG500", "HOMAGKL310", 
    "STREAM1", "STREAM2", "STREAM3", "AKRON1", "AKRON2", "JADE", "NANXING", 
    "SKIPPER1", "SKIPPER2", "SKIPPER3", "SKIPPER4", "SKIPPER5", "ROVER20", "ROVER1", "ROVER2", "BHX1", "BHX2", "FTT", "NESTING", "VITAP" 
]

tecnicos = [
    "WILLIAN DIAZ", "JAIRO ISAZA", "DAIRON MENESES", "LEONARDO ALVAREZ", 
    "BAYRON LOPEZ", "HEBERT CHAPUEL", "ESTEBAN ROSERO", "BRANDON RAMOS", 
    "DAVID PANTOJA", "CARLOS LUGO", "JULIO DAZA", "JHOAN MOTATO"
]

@st.cache_data
def cargar_kardex():
    kardex_file = "KARDEX MTTO.xlsx"
    if os.path.exists(kardex_file):
        try:
            df = pd.read_excel(kardex_file, sheet_name="Saldo", dtype=str, engine='openpyxl')
            dic_repuestos = {}
            for _, row in df.iterrows():
                item_val = row.iloc[0].strip() if pd.notna(row.iloc[0]) else ""
                codigo_val = row.iloc[1].strip() if pd.notna(row.iloc[1]) else ""
                desc_val = row.iloc[2].strip() if pd.notna(row.iloc[2]) else ""
                if item_val.endswith('.0'): item_val = item_val[:-2]
                if codigo_val.endswith('.0'): codigo_val = codigo_val[:-2]
                if item_val and item_val.lower() not in ['nan', 'item']: dic_repuestos[item_val] = desc_val
                if codigo_val and codigo_val.lower() not in ['nan', 'item']: dic_repuestos[codigo_val] = desc_val
            return dic_repuestos
        except Exception:
            return {}
    return {}

diccionario_repuestos = cargar_kardex()

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
st.markdown("Registra fallas con limpieza automática al guardar.")

if not diccionario_repuestos:
    st.warning("⚠️ Nota: No se pudo leer el archivo 'KARDEX MTTO.xlsx'. Asegúrate de subirlo a GitHub.")

# --- CONTADOR DE RESETEO DE FORMULARIO ---
if 'form_id' not in st.session_state:
    st.session_state.form_id = 0

if 'num_repuestos' not in st.session_state:
    st.session_state.num_repuestos = 1

fid = st.session_state.form_id

# --- FORMULARIO ---
with st.form(f"form_reporte_daño_{fid}"):
    fecha = st.date_input("Fecha del reporte", datetime.date.today(), key=f"fecha_{fid}")
    maquina = st.selectbox("Máquina / Equipo", maquinas, key=f"maquina_{fid}")
    
    # Campo SIESA agregado
    siesa = st.text_input("Número de Solicitud / Documento SIESA", key=f"siesa_{fid}")
    
    st.markdown("🕒 **Selección de Tiempos (AM / PM)**")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        str_hora_inicio = st.selectbox("Hora de Inicio del Paro", lista_horas, index=120, key=f"h_ini_{fid}")
    with col_h2:
        str_hora_fin = st.selectbox("Hora de Finalización", lista_horas, index=126, key=f"h_fin_{fid}")
    
    daño = st.text_area("Descripción del Daño / Falla", key=f"dano_{fid}")
    reparacion = st.text_area("Acción de Reparación Realizada", key=f"rep_{fid}")
    
    st.markdown("---")
    st.subheader("Personal de Mantenimiento")
    col3, col4, col5 = st.columns(3)
    with col3:
        tecnico1 = st.selectbox("Técnico 1", [""] + tecnicos, key=f"tec1_{fid}")
    with col4:
        tecnico2 = st.selectbox("Técnico 2", [""] + tecnicos, key=f"tec2_{fid}")
    with col5:
        tecnico3 = st.selectbox("Técnico 3", [""] + tecnicos, key=f"tec3_{fid}")
        
    st.markdown("---")
    st.subheader("📦 Repuestos y Materiales")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        add_rep = st.form_submit_button("➕ Añadir otro repuesto")
    with col_btn2:
        rem_rep = st.form_submit_button("➖ Quitar último repuesto")

    if add_rep:
        st.session_state.num_repuestos += 1
        st.rerun()
    if rem_rep and st.session_state.num_repuestos > 1:
        st.session_state.num_repuestos -= 1
        st.rerun()

    repuestos_lista = []
    cantidades_lista = []
    
    for i in range(st.session_state.num_repuestos):
        st.markdown(f"**Repuesto #{i+1}**")
        col_r1, col_r2 = st.columns([2, 1])
        with col_r1:
            codigo_ingresado = st.text_input(f"Código o Ítem #{i+1}", key=f"cod_{i}_{fid}")
        with col_r2:
            cant = st.text_input(f"Cantidad #{i+1}", value="1", key=f"cant_{i}_{fid}")
        
        if codigo_ingresado.strip():
            codigo_exacto = codigo_ingresado.strip()
            if codigo_exacto in diccionario_repuestos:
                desc_encontrada = diccionario_repuestos[codigo_exacto]
            else:
                match = [v for k, v in diccionario_repuestos.items() if codigo_exacto.lower() in k.lower()]
                desc_encontrada = match[0] if match else "⚠️ Ítem no encontrado"
            
            st.info(f"📄 **Descripción:** {desc_encontrada}")
            repuestos_lista.append(f"[{codigo_exacto}] {desc_encontrada} (Cant: {cant})")
            cantidades_lista.append(cant)
        
        st.markdown("---")

    enviar = st.form_submit_button("💾 Guardar Registro en Google Sheets")
    
    if enviar:
        if not GOOGLE_SCRIPT_URL or "TU_NUEVO_ID_AQUI" in GOOGLE_SCRIPT_URL:
            st.error("Por favor configura la URL válida de Google Apps Script.")
        elif not daño.strip() or not reparacion.strip():
            st.warning("Por favor completa la descripción del daño y la reparación.")
        else:
            def parse_am_pm(t_str, fecha_base):
                t_partes, periodo = t_str.split(" ")
                h_str, m_str = t_partes.split(":")
                h, m = int(h_str), int(m_str)
                if periodo == "PM" and h != 12: h += 12
                if periodo == "AM" and h == 12: h = 0
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
            repuestos_texto = " | ".join(repuestos_lista) if repuestos_lista else ""
            
            cantidad_texto = cantidades_lista[0] if len(cantidades_lista) == 1 else ("Ver detalle" if len(cantidades_lista) > 1 else "")

            payload = {
                "fecha": str(fecha),
                "maquina": maquina,
                "hora_inicio": dt_ini.strftime("%H:%M:%S"),
                "hora_fin": dt_fin.strftime("%H:%M:%S"),
                "tiempo_real": tiempo_real,
                "siesa": siesa,
                "daño": daño,
                "reparacion": reparacion,
                "tecnico1": tecnico1,
                "tecnico2": tecnico2,
                "tecnico3": tecnico3,
                "quien_repara": quien_repara,
                "repuesto": repuestos_texto,
                "cantidad": cantidad_texto
            }
            
            try:
                response = requests.post(
                    GOOGLE_SCRIPT_URL, 
                    data=json.dumps(payload),
                    headers={"Content-Type": "text/plain;charset=utf-8"},
                    timeout=15,
                    allow_redirects=True
                )
                
                if response.status_code == 200:
                    st.session_state.form_id += 1
                    st.session_state.num_repuestos = 1
                    st.toast("✅ ¡Registro guardado y formulario limpiado con éxito!", icon="🎉")
                    st.rerun()
                else:
                    st.error(f"Error en el servidor de Google (Código HTTP: {response.status_code}).")
            except Exception as e:
                st.error(f"Error de conexión: {e}")

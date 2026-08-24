
import streamlit as st
import pandas as pd
import datetime
import requests
import os

st.set_page_config(page_title="Reporte de Daños - Mantenimiento", page_icon="⚙️", layout="centered")

# PEGA AQUÍ TU URL DE GOOGLE APPS SCRIPT (la que termina en /exec)
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzKYQ_9Sx-gRo778XZLF2V_4iUaIs0IHKWTbGHc1q3dEmOI32-gaLLURA4aqE0GSjTg/exec"

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

# Cargar catálogo de repuestos de forma robusta desde KARDEX MTTO.xlsx
@st.cache_data
def cargar_kardex():
    kardex_file = "KARDEX MTTO.xlsx"
    if os.path.exists(kardex_file):
        try:
            df = pd.read_excel(kardex_file, sheet_name="Saldo", engine='openpyxl')
            df = df.dropna(subset=['Desc. Item'])
            
            dic_repuestos = {}
            for _, row in df.iterrows():
                # Limpiar Item principal y código secundario (columna ' ')
                item_val = str(row['Item']).replace('.0', '').strip() if pd.notna(row['Item']) else ""
                codigo_val = str(row[' ']).replace('.0', '').strip() if ' ' in df.columns and pd.notna(row[' ']) else ""
                desc_val = str(row['Desc. Item']).strip()
                
                if item_val and item_val != 'nan':
                    dic_repuestos[item_val] = desc_val
                if codigo_val and codigo_val != 'nan':
                    dic_repuestos[codigo_val] = desc_val
                    
            return dic_repuestos
        except Exception as e:
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
st.markdown("Registra fallas")

if not diccionario_repuestos:
    st.warning("⚠️ Nota: No se detectó o no se pudo leer el archivo 'KARDEX MTTO.xlsx'. Asegúrate de subirlo a GitHub. (Podrás escribir repuestos manualmente si es necesario).")

if 'num_repuestos' not in st.session_state:
    st.session_state.num_repuestos = 1

with st.form("form_reporte_daño"):
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
    st.subheader("Personal de Mantenimiento")
    col3, col4, col5 = st.columns(3)
    with col3:
        tecnico1 = st.selectbox("Técnico 1", [""] + tecnicos)
    with col4:
        tecnico2 = st.selectbox("Técnico 2", [""] + tecnicos)
    with col5:
        tecnico3 = st.selectbox("Técnico 3", [""] + tecnicos)
        
    st.markdown("---")
    st.subheader("📦 Repuestos y Materiales")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.form_submit_button("➕ Añadir otro repuesto"):
            st.session_state.num_repuestos += 1
            st.rerun()
    with col_btn2:
        if st.session_state.num_repuestos > 1 and st.form_submit_button("➖ Quitar último repuesto"):
            st.session_state.num_repuestos -= 1
            st.rerun()

    repuestos_lista = []
    for i in range(st.session_state.num_repuestos):
        col_r1, col_r2, col_r3 = st.columns([2, 2, 1])
        with col_r1:
            codigo_ingresado = st.text_input(f"Código / Item #{i+1}", key=f"cod_{i}")
        with col_r2:
            desc_encontrada = ""
            if codigo_ingresado.strip():
                codigo_limpio = codigo_ingresado.strip().replace('.0', '')
                if codigo_limpio in diccionario_repuestos:
                    desc_encontrada = diccionario_repuestos[codigo_limpio]
                else:
                    # Búsqueda parcial por si escribe parte de la descripción o código
                    match = [v for k, v in diccionario_repuestos.items() if codigo_limpio.lower() in k.lower()]
                    desc_encontrada = match[0] if match else "⚠️ Ítem no encontrado"
            
            st.text_input(f"Descripción #{i+1}", value=desc_encontrada, disabled=True, key=f"desc_{i}")
        with col_r3:
            cant = st.text_input(f"Cant #{i+1}", value="1", key=f"cant_{i}")
        
        if codigo_ingresado.strip():
            repuestos_lista.append(f"[{codigo_ingresado}] {desc_encontrada} (Cant: {cant})")

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
            
            repuestos_texto = " | ".join(repuestos_lista) if repuestos_lista else ""
            
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
                "repuesto": repuestos_texto,
                "cantidad": "Ver detalle" if len(repuestos_lista) > 1 else (cant if 'cant' in locals() and repuestos_lista else "")
            }
            
            try:
                response = requests.post(GOOGLE_SCRIPT_URL, json=payload)
                if response.status_code == 200:
                    st.success("¡Daño registrado y guardado con éxito en Google Sheets!")
                    st.session_state.num_repuestos = 1
                else:
                    st.error("Error al conectar con Google Sheets.")
            except Exception as e:
                st.error(f"Error de conexión: {e}")

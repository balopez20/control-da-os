import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(page_title="Reporte de Daños - Mantenimiento", page_icon="⚙️", layout="centered")

excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
if excel_files:
    EXCEL_FILE = excel_files[0]
else:
    EXCEL_FILE = "daños de mantenimiento.xlsx"

def obtener_opciones():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        
        if "MAQUINAS" in xls.sheet_names:
            df_maq = pd.read_excel(EXCEL_FILE, sheet_name="MAQUINAS").dropna(subset=['MAQUINAS'])
            maquinas = df_maq['MAQUINAS'].tolist()
        else:
            maquinas = ["WNT", "SELCO2", "HOMAG400", "VITAP"]

        if "TECNICOS" in xls.sheet_names:
            df_tec = pd.read_excel(EXCEL_FILE, sheet_name="TECNICOS")
            tecnicos = df_tec['TECNICO1'].dropna().tolist()
        else:
            tecnicos = ["WILLIAN DIAZ", "BAYRON LOPEZ", "DAVID PANTOJA"]
            
        return maquinas, tecnicos
    except Exception as e:
        return ["WNT", "SELCO2"], ["BAYRON LOPEZ", "WILLIAN DIAZ"]

maquinas, tecnicos = obtener_opciones()

# Generar lista de horas AM / PM en intervalos de 5 minutos
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
st.markdown("Selecciona la máquina, las horas de inicio/fin y registra la falla fácilmente.")

with st.form("form_reporte_daño", clear_on_submit=True):
    fecha = st.date_input("Fecha del reporte", datetime.date.today())
    
    maquina = st.selectbox("Máquina / Equipo", maquinas)
    
    st.markdown("🕒 **Selección de Tiempos (AM / PM)**")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        str_hora_inicio = st.selectbox("Hora de Inicio del Paro", lista_horas, index=120) # Por defecto alrededor de las 10:00 AM
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
    
    enviar = st.form_submit_button("💾 Guardar Registro en el Excel")
    
    if enviar:
        if not daño.strip() or not reparacion.strip():
            st.warning("Por favor completa la descripción del daño y la reparación.")
        else:
            # Función para convertir formato AM/PM a objeto time y datetime
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
            
            if dt_fin < dt_ini: # Por si cruza la medianoche
                dt_fin += datetime.timedelta(days=1)
            
            diferencia = dt_fin - dt_ini
            horas = diferencia.seconds // 3600
            minutos = (diferencia.seconds % 3600) // 60
            tiempo_real = f"{horas:02d}:{minutos:02d}:00"
            
            lista_repara = [t for t in [tecnico1, tecnico2, tecnico3] if t != ""]
            quien_repara = ", ".join(lista_repara) if lista_repara else ""
            
            nuevo_registro = {
                "MAQUINAS": maquina,
                "AREA": "", # Campo de área omitido por solicitud
                "HORA INICIO": dt_ini.strftime("%H:%M:%S"),
                "HORA FIN": dt_fin.strftime("%H:%M:%S"),
                "TIEMPO REAL": tiempo_real,
                "NUMERO DE SOLICITUD": "",
                "DAÑO": daño,
                "REPARACION": reparacion,
                "TECNICO 1": tecnico1,
                "TECNICO2": tecnico2,
                "TECNICO3": tecnico3,
                "QUIEN REPARA": quien_repara,
                "REPUESTOS ITEM": repuesto,
                "CANTIDAD": cantidad if repuesto else "",
                "DESCRIPCION": daño
            }
            
            try:
                con_excel = pd.ExcelFile(EXCEL_FILE)
                hoja_objetivo = "Agosto" if "Agosto" in con_excel.sheet_names else con_excel.sheet_names[0]
                
                df_existente = pd.read_excel(EXCEL_FILE, sheet_name=hoja_objetivo)
                df_nuevo = pd.concat([df_existente, pd.DataFrame([nuevo_registro])], ignore_index=True)
                
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
                    df_nuevo.to_excel(writer, sheet_name=hoja_objetivo, index=False)
                    for sheet in con_excel.sheet_names:
                        if sheet != hoja_objetivo:
                            pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
                
                st.success("¡Daño registrado y guardado con éxito en el Excel!")
            except Exception as e:
                st.error(f"Error al guardar en el archivo Excel: {e}")

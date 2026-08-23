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
            areas = df_maq['AREA'].dropna().unique().tolist()
        else:
            maquinas = ["WNT", "SELCO2", "HOMAG400", "VITAP"]
            areas = ["CORTE", "ENCHAPE", "MAQUINADO"]

        if "TECNICOS" in xls.sheet_names:
            df_tec = pd.read_excel(EXCEL_FILE, sheet_name="TECNICOS")
            tecnicos = df_tec['TECNICO1'].dropna().tolist()
        else:
            tecnicos = ["WILLIAN DIAZ", "BAYRON LOPEZ", "DAVID PANTOJA"]
            
        return maquinas, areas, tecnicos
    except Exception as e:
        return ["WNT", "SELCO2"], ["CORTE"], ["BAYRON LOPEZ", "WILLIAN DIAZ"]

maquinas, areas, tecnicos = obtener_opciones()

st.title("📱 Reporte Diario de Daños")
st.markdown("Registra las fallas de planta de forma rápida desde el celular.")

with st.form("form_reporte_daño", clear_form=True):
    fecha = st.date_input("Fecha del reporte", datetime.date.today())
    
    col1, col2 = st.columns(2)
    with col1:
        maquina = st.selectbox("Máquina / Equipo", maquinas)
    with col2:
        area = st.selectbox("Área", areas)
        
    hora_inicio = st.time_input("Hora de Inicio del Paro", datetime.datetime.now().time())
    hora_fin = st.time_input("Hora de Finalización", datetime.datetime.now().time())
    
    daño = st.text_area("Descripción del Daño / Falla")
    reparacion = st.text_area("Acción de Reparación Realizada")
    
    st.markdown("---")
    st.subheader("Personal y Repuestos")
    
    col3, col4 = st.columns(2)
    with col3:
        tecnico1 = st.selectbox("Técnico 1 (Responsable)", tecnicos)
    with col4:
        tecnico2 = st.selectbox("Técnico 2 (Apoyo)", ["Ninguno"] + tecnicos)
        
    repuesto = st.text_input("Repuesto utilizado (Opcional)")
    cantidad = st.number_input("Cantidad", min_value=0, step=1)
    
    enviar = st.form_submit_button("💾 Guardar Registro en el Excel")
    
    if enviar:
        if not daño.strip() or not reparacion.strip():
            st.warning("Por favor completa la descripción del daño y la reparación.")
        else:
            h_ini = datetime.datetime.combine(fecha, hora_inicio)
            h_fin = datetime.datetime.combine(fecha, hora_fin)
            if h_fin < h_ini:
                h_fin += datetime.timedelta(days=1)
            tiempo_real = str(h_fin - h_ini)
            
            nuevo_registro = {
                "MAQUINAS": maquina,
                "AREA": area,
                "HORA INICIO": str(hora_inicio),
                "HORA FIN": str(hora_fin),
                "TIEMPO REAL": tiempo_real,
                "NUMERO DE SOLICITUD": "",
                "DAÑO": daño,
                "REPARACION": reparacion,
                "TECNICO 1": tecnico1,
                "TECNICO2": tecnico2 if tecnico2 != "Ninguno" else "",
                "TECNICO3": "",
                "QUIEN REPARA": f"{tecnico1}, {tecnico2}" if tecnico2 != "Ninguno" else tecnico1,
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
                
                st.success("¡Daño registrado y guardado en el Excel con éxito!")
            except Exception as e:
                st.error(f"Error al guardar en el archivo Excel: {e}")
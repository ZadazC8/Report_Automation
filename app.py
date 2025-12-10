import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import date
import io
import os
import re
import base64

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    layout="wide", 
    page_title="Generador de Informe", 
    page_icon="assets/logo_empresa.png"
)

# --- FUNCIÓN AUXILIAR PARA IMAGEN ---
def img_to_base64(file_path):
    with open(file_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

# --- 2. DICCIONARIO DE TRADUCCIONES ---
TRADUCCIONES = {
    'ES': {
        # Etiquetas (Labels)
        'titulo_pdf': 'INFORME DE ENSAYO DE LABORATORIO',
        'codigo': 'CÓDIGO', 'version': 'VERSIÓN', 'fecha': 'FECHA', 'pagina': 'Página',
        'target': 'OBJETIVO', 'cliente': 'CLIENTE', 'orden': 'PEDIDO/ORDEN',
        'direccion': 'DIRECCIÓN', 
        'procedimiento': 'PROCEDIMIENTO',
        'ensayo_realizado': 'ENSAYO REALIZADO',
        'realizado_por': 'REALIZADO POR', 'fecha_test': 'FECHA ENSAYO',
        'fecha_emision': 'FECHA EMISIÓN',
        
        # VALORES DEL CONTENIDO (TRADUCIBLES)
        'val_procedimiento': 'Procedimiento Interno No. EFU-P-032',
        'val_target': 'Prueba de resistividad utilizando miliohmímetro Harris.',
        'test': 'Resistencia Eléctrica', # <--- AQUI ESTABA EL PROBLEMA
        
        # Resto del contenido
        'disclaimer': 'Este informe no podrá ser reproducido total o parcialmente sin la aprobación por escrito del Laboratorio de Eléctricos Internacional.',
        'indice': 'ÍNDICE DE TABLAS',
        'tabla1_titulo': 'DATOS DE PLACA',
        't1_id': 'IDENTIFICACIÓN', 't1_placa': 'PLACA DE CARACTERÍSTICAS',
        't1_tipo': 'TIPO (CABEZA REMOVIBLE)', 't1_corriente': 'CORRIENTE NOMINAL (A)',
        't1_voltaje': 'VOLTAJE NOMINAL (kV)', 't1_cant': 'CANTIDAD',
        't1_total': 'CANTIDAD TOTAL DE UNIDADES',
        'nota_muestras': '* Nota: Las muestras fueron suministradas por el cliente.',
        'tabla2_titulo': 'EQUIPO UTILIZADO',
        't2_item': 'Ítem', 't2_equipo': 'EQUIPO', 't2_marca': 'MARCA',
        't2_modelo': 'MODELO', 't2_serial': 'NÚMERO DE SERIE', 't2_cert': 'CERTIFICADO CALIBRACIÓN',
        'no_equipos': 'No se cargaron datos de equipos',
        'tabla3_titulo': 'RESULTADOS GENERALES DEL ENSAYO',
        'ref': 'REF', 'cant': 'CANT', 'rango': 'RANGO (mΩ)',
        'promedio': 'PROMEDIO', 'max': 'MÁX', 'min': 'MÍN',
        'posibles_resultados': 'POSIBLES RESULTADOS',
        'txt_pass': 'PASS (CONFORME): Los resultados satisfacen los requisitos del procedimiento interno',
        'txt_fail': 'FAIL (NO CONFORME): Los resultados no cumplen con los requisitos del procedimiento interno',
        'condiciones': 'CONDICIONES DEL ENSAYO - RESISTENCIA ELÉCTRICA',
        'temp': 'TEMPERATURA AMBIENTE PROMEDIO', 'hum': 'HUMEDAD RELATIVA',
        'detalles_titulo': 'RESULTADOS DETALLADOS',
        'titulo_tecnico_tabla': 'MEDICIÓN DE RESISTENCIA ELÉCTRICA - FUSIBLE TIPO T',
        'muestra': 'MUESTRA', 'lectura': 'LECTURA (mΩ)', 'resultado': 'RESULTADO',
        'nota_final': 'Los datos y resultados contenidos en este informe corresponden únicamente a la muestra suministrada por el cliente al laboratorio como se indica en el párrafo 1. De los resultados obtenidos se concluye que las muestras ensayadas pasaron las pruebas descritas en el',
        'aprobado_por': 'Aprobado Por',
        'borrador': 'INFORME BORRADOR', 'no_valido': 'Sin validez oficial',
        'fin_doc': '*** FIN DEL DOCUMENTO ***'
    },
    'EN': {
        # Labels
        'titulo_pdf': 'LABORATORY TEST REPORT',
        'codigo': 'CODE', 'version': 'VERSION', 'fecha': 'DATE', 'pagina': 'Page',
        'target': 'TARGET', 'Ensayo' : 'TEST PERFORMED','cliente': 'CLIENT', 'orden': 'TRACE PRODUCT',
        'direccion': 'ADDRESS', 
        'procedimiento': 'PROCEDURE',
        'ensayo_realizado': 'TEST PERFORMED',
        'realizado_por': 'TESTED BY', 'fecha_test': 'TEST DATE',
        'fecha_emision': 'ISSUE DATE',
        
        # VALUES (English)
        'val_procedimiento': 'Internal procedure No. EFU-P-032',
        'val_target': 'Resistivity testing using miliohmimeter Harris.',
        'test': 'Electric Resistance',
        
        # Content
        'disclaimer': 'This report shall not be reproduced in full, without the written approval of Electricos Internacional Laboratory.',
        'indice': 'TABLE INDEX',
        'tabla1_titulo': 'FUSE LINKS NAMEPLATE DATA',
        't1_id': 'IDENTIFICACIÓN', 't1_placa': 'FUSE LINKS NAMEPLATE',
        't1_tipo': 'TYPE (REMOVABLE HEAD)', 't1_corriente': 'NOMINAL CURRENT (A)',
        't1_voltaje': 'NOMINAL VOLTAGE (kV)', 't1_cant': 'QUANTITY',
        't1_total': 'TOTAL NUMBER OF UNITS',
        'nota_muestras': '* Note: Samples were provided by the customer.',
        'tabla2_titulo': 'EQUIPMENT USED',
        't2_item': 'Item', 't2_equipo': 'EQUIPMENT', 't2_marca': 'BRAND',
        't2_modelo': 'MODEL', 't2_serial': 'SERIAL NUMBER', 't2_cert': 'CALIBRATION CERTIFICATE',
        'no_equipos': 'No equipment data loaded',
        'tabla3_titulo': 'OVERALL TESTS RESULTS',
        'ref': 'RATING', 'cant': 'QTY', 'rango': 'RANGE (mΩ)',
        'promedio': 'AVERAGE', 'max': 'MAX', 'min': 'MIN',
        'posibles_resultados': 'POSSIBLE RESULTS',
        'txt_pass': 'PASS: The results satisfy the requirements of the internal procedure',
        'txt_fail': 'FAIL: The results are not in accordance with the requirements of the internal procedure',
        'condiciones': 'TEST CONDITIONS - ELECTRIC RESISTANCE',
        'temp': 'AVERAGE AMBIENT TEMPERATURE', 'hum': 'RELATIVE HUMIDITY',
        'detalles_titulo': 'DETAILED RESULTS',
        'titulo_tecnico_tabla': 'MEASUREMENT OF ELECTRICAL RESISTANCE - FUSE LINK T',
        'muestra': 'SAMPLE ID', 'lectura': 'READING (mΩ)', 'resultado': 'RESULT',
        'nota_final': 'The data and results contained in this report, only corresponds to the sample supplied by the client to the laboratory as listed in paragraph 1 (Description of samples). From the results obtained it is concluded that the samples tested passed the tests described on the',
        'aprobado_por': 'Approved By',
        'borrador': 'DRAFT REPORT', 'no_valido': 'Not valid for official use',
        'fin_doc': '*** END OF THIS DOCUMENT ***'
    }
}

# --- 3. GENERADOR DE PLANTILLA EXCEL ---
def generar_plantilla():
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    # 1. Hoja Ejemplo
    df_ejemplo = pd.DataFrame({
        'ID_Muestra': ['LABM-1', 'LABM-2', 'LABM-3'],
        'Lectura_mOhm': [1.79, 1.73, 1.78],
        'Tipo_Fusible': ['T-SLOW ELEMENT (FL6T30)', 'T-SLOW ELEMENT (FL6T30)', 'T-SLOW ELEMENT (FL6T30)'],
        'Voltaje_Nominal': [15, 15, 15]
    })
    df_ejemplo.to_excel(writer, sheet_name='30A (Ejemplo)', index=False)
    
    # 2. Hoja Equipos
    df_equipos = pd.DataFrame({
        'Item': [1, 2],
        'EQUIPMENT': ['INDUSTRIAL RESISTANCE TESTER', 'CALIBRATOR'],
        'BRAND': ['HARRIS IRT', 'FLUKE'],
        'MODEL': ['5060-06XR', '754'],
        'SERIAL NUMBER': ['0801010', '9876543'],
        'CALIBRATION CERTIFICATE': ['FLORIDA No. 131773', 'METROLOGY No. 555']
    })
    df_equipos.to_excel(writer, sheet_name='EQUIPMENT', index=False)
    
    writer.close()
    return output.getvalue()

# --- 4. ESTILOS CSS ---
st.markdown("""
<style>
    :root { --primary: #002f6c; --secondary: #004a9c; --accent: #ff6b35; --bg: #f8f9fa; }
    .main-header {
        background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
        padding: 1rem 2rem; border-radius: 15px; margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0, 47, 108, 0.2); color: white;
        display: flex; align-items: center; gap: 20px;
    }
    .header-logo { height: 70px; width: auto; background-color: white; padding: 5px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
    .header-text h1 { font-size: 2.2rem; font-weight: 800; margin: 0; color: white !important; line-height: 1.2; }
    .header-text p { font-size: 1rem; margin: 0; opacity: 0.9; font-weight: 300; }
    [data-testid="stMetricValue"] { color: var(--primary); font-weight: 700; }
    .stButton button { background-color: var(--primary) !important; color: white !important; transition: 0.3s; }
    .stButton button:hover { background-color: var(--accent) !important; transform: translateY(-2px); }
    [data-testid="stFileUploader"] { border: 1px dashed #004a9c; border-radius: 8px; background-color: #f0f2f6; }
    [data-testid="stDataFrame"] { border: 1px solid #e8f0f8; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# --- 5. LÓGICA DE NEGOCIO ---
RANGOS_TIPO_T = {
    "6A": {"min": 17.2, "max": 22.1}, "8A": {"min": 12.4, "max": 16.0},
    "10A": {"min": 8.1, "max": 10.4}, "12A": {"min": 6.2, "max": 8.0},
    "15A": {"min": 4.1, "max": 5.2}, "20A": {"min": 3.5, "max": 4.4},
    "25A": {"min": 2.6, "max": 3.3}, "30A": {"min": 1.6, "max": 2.0},
    "40A": {"min": 1.3, "max": 1.7}, "50A": {"min": 1.1, "max": 1.4},
    "65A": {"min": 0.7, "max": 0.9}, "80A": {"min": 0.5, "max": 0.7},
    "100A": {"min": 0.3, "max": 0.5}, "140A": {"min": 0.2, "max": 0.3},
    "200A": {"min": 0.1, "max": 0.2}
}
T_EVAL = {'pass': 'PASS', 'fail': 'FAIL'}

def obtener_limites(nombre_hoja):
    nombre = nombre_hoja.upper()
    for rating, limites in RANGOS_TIPO_T.items():
        if rating in nombre:
            corriente_num = re.findall(r'\d+', rating)[0] 
            return limites['min'], limites['max'], rating, corriente_num
    return 0, 9999, nombre_hoja, "N/A"

# --- 6. BARRA LATERAL ---
with st.sidebar:
    st.header("Plantilla")
    plantilla_excel = generar_plantilla()
    st.download_button("Descargar Formato Excel", data=plantilla_excel, file_name="Plantilla_Ensayo_Laboratorio.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key='btn_plantilla', use_container_width=True)
    st.divider()
    st.header("Cargar Datos")
    uploaded_file = st.file_uploader("Excel con Resultados", type=["xlsx"])
    st.divider()
    st.markdown("### Configuración")
    idioma_seleccionado = st.radio("Idioma del Reporte PDF", ["Español", "Inglés"], index=1)
    lang_code = 'ES' if idioma_seleccionado == "Español" else 'EN'
    st.divider()
    st.markdown("### Aprobación")
    PIN_SECRET = "LAB123" 
    password = st.text_input("PIN de Firma", type="password")
    autorizado = False
    nombre_aprobador = "" 
    cargo_aprobador = "" 
    if password == PIN_SECRET:
        autorizado = True
        st.success("Firma HABILITADA")
        nombre_aprobador = st.text_input("Nombre Aprobador", value="")
        cargo_aprobador = st.text_input("Cargo Aprobador", value="")
    elif password != "":
        st.error("PIN Incorrecto")

# --- 7. INTERFAZ PRINCIPAL ---
logo_path = "assets/logo_empresa.png"
img_tag = ""
if os.path.exists(logo_path):
    img_b64 = img_to_base64(logo_path)
    img_tag = f'<img src="data:image/png;base64,{img_b64}" class="header-logo">'

st.markdown(f"""
<div class="main-header">
    {img_tag}
    <div class="header-text">
        <h1>Generador de Informes - Resistividad de Fusibles</h1>
        <p>Laboratorio Eléctricos Internacional S.A.S</p>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("Datos del Pedido", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    with c1: consecutivo = st.text_input("Consecutivo Informe", value="")
    with c2: orden = st.text_input("Número de Pedido", value="")
    with c3: fecha_prueba = st.date_input("Fecha Prueba", date.today())
    with c4: fecha_emision = st.date_input("Fecha Emisión", date.today())
    c5, c6 = st.columns([2, 2])
    with c5: cliente = st.text_input("Cliente", value="")
    with c6: direccion = st.text_input("Dirección", value="")
    c7, c8 = st.columns(2)
    with c7: operador = st.text_input("Realizado por (Operador)", value="")
    with c8: st.info(f"El informe se generará en: **{idioma_seleccionado}**")

# --- 8. PROCESAMIENTO ---
if uploaded_file is not None:
    st.divider()
    st.subheader("Vista Previa de Resultados")
    xls = pd.ExcelFile(uploaded_file)
    hojas = [h for h in xls.sheet_names if h != 'INFO' and h != 'EQUIPMENT']
    resultados = []
    total_muestras_global = 0
    total_fallas = 0
    
    lista_equipos = []
    try:
        df_equipos = pd.read_excel(uploaded_file, sheet_name='EQUIPMENT')
        df_equipos = df_equipos.fillna("-") 
        lista_equipos = df_equipos.to_dict('records')
    except Exception:
        st.warning("No se encontró la hoja 'EQUIPMENT'.")

    tabs = st.tabs(hojas)
    for i, hoja in enumerate(hojas):
        with tabs[i]:
            lim_min, lim_max, nombre_ref, corriente_nominal = obtener_limites(hoja)
            df = pd.read_excel(uploaded_file, sheet_name=hoja)
            req_cols = ['ID_Muestra', 'Lectura_mOhm']
            if not all(col in df.columns for col in req_cols):
                st.error(f"Faltan columnas en '{hoja}'. Requerido: {req_cols}"); continue
            df = df.dropna(subset=['Lectura_mOhm'])
            tipo_fusible = df['Tipo_Fusible'].iloc[0] if 'Tipo_Fusible' in df.columns else "Generic"
            voltaje_nominal = df['Voltaje_Nominal'].iloc[0] if 'Voltaje_Nominal' in df.columns else "15"
            id_txt = f"{df['ID_Muestra'].iloc[0]} - {df['ID_Muestra'].iloc[-1]}"
            
            lista_datos = []
            fallas_locales = 0
            for _, row in df.iterrows():
                val = row['Lectura_mOhm']
                stado = T_EVAL['pass'] if (lim_min <= val <= lim_max) else T_EVAL['fail']
                if stado == T_EVAL['fail']: fallas_locales += 1
                lista_datos.append({'id': row['ID_Muestra'], 'valor': val, 'estado': stado})
            total_fallas += fallas_locales
            total_muestras_global += len(df)
            
            stats = {
                'nombre': nombre_ref, 'corriente_nominal': corriente_nominal,
                'tipo': tipo_fusible, 'voltaje': voltaje_nominal, 'identificacion': id_txt,
                'cantidad': len(df), 'rango_txt': f"{lim_min} - {lim_max}",
                'promedio': df['Lectura_mOhm'].mean(), 'maximo': df['Lectura_mOhm'].max(), 
                'minimo': df['Lectura_mOhm'].min(), 'datos_detalle': lista_datos
            }
            resultados.append(stats)
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ref", nombre_ref)
            m2.metric("Tipo", tipo_fusible)
            m3.metric("Voltaje", f"{voltaje_nominal} kV")
            if fallas_locales > 0: m4.metric("Estado", f"{fallas_locales} FALLAS", delta_color="inverse")
            else: m4.metric("Estado", "OK", delta_color="normal")
            st.dataframe(pd.DataFrame(lista_datos).style.applymap(lambda v: f'background-color: {"#d4edda" if v=="PASS" else "#f8d7da"}; font-weight: bold;', subset=['estado']), use_container_width=True, height=200)

    st.divider()
    c_res1, c_res2 = st.columns([3, 2])
    with c_res1:
        if total_fallas > 0: st.warning(f"Atención: Hay {total_fallas} fallas en total.")
        else: st.success("Todo validado correctamente.")
    
    with c_res2:
        btn_label = f"GENERAR PDF ({idioma_seleccionado.upper()})" + (" FIRMADO" if autorizado else " BORRADOR")
        btn_type = "primary" if autorizado else "secondary"
        
        if st.button(btn_label, type=btn_type, key='btn_generar_pdf', use_container_width=True):
            with st.spinner("Generando documento..."):
                # SELECCIONAMOS DICCIONARIO
                textos_pdf = TRADUCCIONES[lang_code]
                
                info = {
                    'Consecutivo': consecutivo, 'Cliente': cliente, 'Direccion': direccion,
                    'Orden': orden, 'FechaPrueba': fecha_prueba.strftime("%Y-%m-%d"),
                    'FechaEmision': fecha_emision.strftime("%Y-%m-%d"),
                    'Procedimiento': textos_pdf['val_procedimiento'],
                    'Target': textos_pdf['val_target'],
                    'Ensayo': textos_pdf['test'],
                    'TotalMuestras': total_muestras_global,
                    'Operador': operador, 'Autorizado': autorizado,
                    'Aprobador': nombre_aprobador, 'CargoAprobador': cargo_aprobador
                }
                
                env = Environment(loader=FileSystemLoader('templates'))
                template = env.get_template('informe_v1.html')
                html = template.render(
                    info=info, 
                    datos=resultados, 
                    equipos=lista_equipos, 
                    t=textos_pdf, 
                    hoy=date.today()
                )
                
                pdf = io.BytesIO()
                HTML(string=html, base_url='.').write_pdf(pdf)
                file_clean = consecutivo.replace('/', '-').replace('\\', '-') if consecutivo else "Informe"
                st.download_button("DESCARGAR PDF", data=pdf, file_name=f"{file_clean}.pdf", mime="application/pdf", key='btn_descargar_pdf', use_container_width=True)

else:
    st.info("Por favor, carga el archivo Excel desde la barra lateral izquierda para comenzar.")
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from datetime import datetime
import os
from io import StringIO

# ================= CONFIG =================
st.set_page_config(
    page_title="Automatizador de Cartas",
    page_icon="⚡",
    layout="wide"
)

# CONFIGURACIÓN DE REINICIO
if 'count_reset' not in st.session_state:
    st.session_state.count_reset = 0

if 'carta_generada' not in st.session_state:
    st.session_state.carta_generada = False

if 'output_path' not in st.session_state:
    st.session_state.output_path = None

if 'vista_previa_html' not in st.session_state:
    st.session_state.vista_previa_html = None

os.makedirs("templates", exist_ok=True)
os.makedirs("output", exist_ok=True)

# ================= GERENTES POR ZONA =================
GERENTES = {
    "Norte": "Christian Alberto Gómez Díaz",
    "Centro": "Alex Andrés González Villablanca",
    "Sur": "Christian Enrique Araya Silva"
}

# ================= TIPOS DE CARTAS =================
TIPOS_CARTAS = {
    "error_lectura": "Error de Lectura",
    "aumento_consumo": "Aumento de Consumo",
    "cobro_verificacion": "Cobro Verificación Medidor",
    "peak_consumo": "Peak de Consumo",
    "cobro_indebido": "Cobro Indebido"
}

# ================= CANALES CON GÉNERO =================
CANALES_MASCULINO = ["Call Center", "Correo Electrónico", "Portal SEC"]
CANALES_FEMENINO = ["Oficina Comercial", "Página Web", "App CGE 1Click"]

def formatear_monto(valor):
    """Formatea monto con $ y puntos de miles"""
    try:
        valor_limpio = str(valor).replace('$', '').replace('.', '').replace(',', '').strip()
        if valor_limpio:
            numero = int(valor_limpio)
            return f"${numero:,}".replace(',', '.')
        return valor
    except:
        return valor

# ================= HEADER =================
st.markdown("""
<div style='background:linear-gradient(90deg,#1e40af,#0ea5e9);
padding:25px;border-radius:12px;margin-bottom:25px;box-shadow:0 4px 6px rgba(0,0,0,0.1)'>
<h1 style='color:white;margin:0;font-size:2.2em'>⚡ AUTOMATIZADOR DE CARTAS </h1>
<p style='color:#e0f2fe;margin:8px 0 0 0;font-size:1.1em'>Sistema de Generación Automática de Respuestas</p>
</div>
""", unsafe_allow_html=True)

# ================= SELECTOR TIPO DE CARTA =================
st.markdown("### 📋 SELECCIONAR TIPO DE CARTA")
tipo_carta = st.selectbox(
    "Tipo de respuesta:",
    options=list(TIPOS_CARTAS.keys()),
    format_func=lambda x: TIPOS_CARTAS[x],
    key=f"tipo_carta_{st.session_state.count_reset}"
)

# ================= BOTÓN REINICIAR =================
if st.button("🔄 REINICIAR FORMULARIO"):
    # Sumamos 1 al contador para forzar a Streamlit a olvidar los textos escritos
    st.session_state.count_reset += 1 
    
    # Limpiamos los estados de la carta
    st.session_state.carta_generada = False
    st.session_state.output_path = None
    st.session_state.vista_previa_html = None
    st.rerun()

# ================= FORMULARIO =================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📋 DATOS DEL CLIENTE")
    
    comuna = st.text_input(
        "Comuna:",
        placeholder="Valparaiso",
        key=f"comuna_{st.session_state.count_reset}",
        help="Ciudad/comuna donde se emite la carta"
    )
    
    tratamiento = st.selectbox(
        "Tratamiento:", 
        ["", "Señor", "Señora"],
        key=f"tratamiento_{st.session_state.count_reset}"
    )
    
    nombre_cliente = st.text_input(
        "Nombre y apellido completo:",
        placeholder="Eduardo López",
        key=f"nombre_{st.session_state.count_reset}"
    )
    
    direccion = st.text_input(
        "Dirección:",
        placeholder="Prat 725",
        key=f"direccion_{st.session_state.count_reset}"
    )
    
    numero_cliente = st.text_input(
        "Número cliente:",
        placeholder="6255126",
        key=f"num_cliente_{st.session_state.count_reset}",
        autocomplete="off"
    )

with col2:
    st.markdown("### 📄 DATOS DEL RECLAMO")
    
    gr_numero = st.text_input(
        "N° GR (Número de Reclamo):",
        placeholder="15624563",
        key=f"gr_{st.session_state.count_reset}",
        help="Este número se usará en DGR y en Ref.: Reclamo N°",
        autocomplete="off"
    )
    
    zona = st.selectbox(
        "Firma - Zona Geográfica:",
        ["Norte", "Centro", "Sur"],
        help="Seleccione la zona para el gerente comercial correspondiente",
        key=f"zona_{st.session_state.count_reset}"
    )
    
    st.info(f"👤 Gerente Comercial: {GERENTES[zona]}")
    
    canal_ingreso = st.selectbox(
        "Canal de ingreso del reclamo:",
        [
            "Oficina Comercial",
            "WhatsApp",
            "App CGE 1Click",
            "Call Center",
            "Correo Electrónico",
            "Página Web",
            "Portal SEC"
        ],
        key=f"canal_{st.session_state.count_reset}"
    )

st.markdown("---")

# ================= DATOS ESPECÍFICOS POR TIPO DE CARTA =================
if tipo_carta == "error_lectura":
    with st.expander("📊 DATOS ADICIONALES DE FACTURACIÓN (Opcional)"):
        col_a, col_b, col_c = st.columns(3)
        
        with col_a:
            fecha_boleta = st.text_input(
                "Fecha boleta:", 
                placeholder="15/12/2025",
                key=f"fecha_boleta_{st.session_state.count_reset}"
            )
            tipo_doc = st.selectbox(
                "Tipo documento:", 
                ["boleta", "factura"],
                key=f"tipo_doc_{st.session_state.count_reset}"
            )
        
        with col_b:
            consumo_kwh = st.text_input(
                "Consumo kWh:", 
                placeholder="350",
                key=f"consumo_kwh_{st.session_state.count_reset}"
            )
            monto_boleta_input = st.text_input(
                "Monto:", 
                placeholder="45000",
                key=f"monto_{st.session_state.count_reset}"
            )
            monto_boleta = formatear_monto(monto_boleta_input) if monto_boleta_input else ""
            if monto_boleta:
                st.caption(f"💰 Formateado: {monto_boleta}")
        
        with col_c:
            dia_inicio = st.text_input(
                "Día inicio lectura:", 
                placeholder="06",
                key=f"dia_inicio_{st.session_state.count_reset}"
            )
            dia_fin = st.text_input(
                "Día fin lectura:", 
                placeholder="12",
                key=f"dia_fin_{st.session_state.count_reset}"
            )
            rango_lectura = f"{dia_inicio} y {dia_fin}" if dia_inicio and dia_fin else ""

# ================= TABLA EXCEL =================
with st.expander("📎 TABLA DE DATOS (Opcional - se agrega al final)"):
    st.info("💡 Copia y pega tabla desde Excel usando TAB como separador")
    tabla_excel = st.text_area(
        "Pegar datos aquí:",
        height=150,
        placeholder="Período\tConsumo\tMonto\nEne-2025\t150\t$45.000",
        key=f"tabla_excel_{st.session_state.count_reset}"
    )
    
    df = None
    if tabla_excel:
        try:
            df = pd.read_csv(StringIO(tabla_excel), sep="\t")
            st.success(f"✅ {len(df)} filas cargadas")
            st.dataframe(df, use_container_width=True)
        except:
            st.error("⚠️ Formato incorrecto. Asegúrate de separar con TAB")

st.markdown("---")

# ================= BOTONES GENERAR Y DESCARGAR =================
col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1, 1.2, 1.2, 1])

with col_btn2:
    generar = st.button(
        "📝 GENERAR CARTA", 
        type="primary", 
        use_container_width=True,
        key=f"btn_generar_{st.session_state.count_reset}"
    )

with col_btn3:
    # El botón de descarga siempre revisa si hay algo generado en la memoria
    if st.session_state.get('carta_generada') and st.session_state.get('output_path'):
        if os.path.exists(st.session_state.output_path):
            with open(st.session_state.output_path, "rb") as f:
                st.download_button(
                    label="📥 DESCARGAR CARTA",
                    data=f,
                    file_name=os.path.basename(st.session_state.output_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key=f"btn_descarga_{st.session_state.count_reset}"
                )

if generar:
    campos_requeridos = {
        "Nombre cliente": nombre_cliente,
        "N° GR": gr_numero,
        "Dirección": direccion,
        "Comuna": comuna,
        "Número cliente": numero_cliente,
        "Tratamiento": tratamiento
    }
    
    faltantes = [k for k, v in campos_requeridos.items() if not v]
    
    if faltantes:
        st.error(f"❌ Faltan campos obligatorios: {', '.join(faltantes)}")
    else:
        template_path = f"templates/{tipo_carta}.docx"
        
        if not os.path.exists(template_path):
            st.error(f"❌ No se encontró el template: {template_path}")
            st.info(f"📝 Copia tu carta Word a la carpeta 'templates/' como '{tipo_carta}.docx'")
        else:
            with st.spinner("⚡ Generando carta..."):
                try:
                    doc = Document(template_path)
                    
                    hoy = datetime.now()
                    meses = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio",
                            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                    
                    if canal_ingreso == "WhatsApp":
                        texto_canal = "WhatsApp"
                    elif canal_ingreso in CANALES_MASCULINO:
                        texto_canal = f"nuestro {canal_ingreso}"
                    else:
                        texto_canal = f"nuestra {canal_ingreso}"
                    
                    estimado = "Estimado" if tratamiento == "Señor" else "Estimada"
                    primer_nombre = nombre_cliente.split()[0]
                    
                    # REEMPLAZOS - Los más específicos primero
                    reemplazos = {
                        # Fecha
                        "[Comuna]": comuna,
                        "[día]": str(hoy.day),
                        "[mes]": meses[hoy.month],
                        "202[X]": str(hoy.year),
                        "[202X]": str(hoy.year),
                        
                        # DGR - Encabezado
                        "DGR N.º [XXXXXXX] /[202X]": f"DGR N° {gr_numero} /{hoy.year}",
                        "DGR N.º XXXXXXX /[202X]": f"DGR N° {gr_numero} /{hoy.year}",
                        "DGR N.º [XXXXXXX]/[202X]": f"DGR N° {gr_numero}/{hoy.year}",
                        "DGR N.º XXXXXXX/[202X]": f"DGR N° {gr_numero}/{hoy.year}",
                        
                        # REF - Esta es la parte que te fallaba
                        "Ref.: Reclamo N° 15965848": f"Ref.: Reclamo N° {gr_numero}",
                        "Ref.: Reclamo N° [XXXXXXX]": f"Ref.: Reclamo N° {gr_numero}",
                        "Ref.: Reclamo N° XXXXXXX": f"Ref.: Reclamo N° {gr_numero}",
                        
                        # Cliente y Datos
                        "Número de cliente: 15965848": f"Número de cliente: {numero_cliente}",
                        "Número de cliente: [XXXXXXX]": f"Número de cliente: {numero_cliente}",
                        "Número de cliente: XXXXXXX": f"Número de cliente: {numero_cliente}",
                        "«[XXXXXXX]»": numero_cliente,
                        
                        # Genéricos de respaldo
                        "[XXXXXXX]": gr_numero,
                        "XXXXXXX": gr_numero,
                        
                        # Cliente info
                        "[Señor(a)]": tratamiento,
                        "Señ[or(a)]": tratamiento,
                        "[«Nombre y apellido reclamante»]": nombre_cliente,
                        "[Nombre y apellido reclamante]": nombre_cliente,
                        "[«Dirección»]": direccion,
                        "[Dirección]": direccion,
                        "[«Comuna»]": comuna,
                        
                        # Saludo
                        "[Estimado(a) Nombre,]": f"{estimado} {primer_nombre},",
                        "Estimado(a) [Nombre,]": f"{estimado} {primer_nombre},",
                        "[Nombre,]": f"{primer_nombre},",
                        
                        # Canal
                        "[(Ej: nuestra Oficina Comercial / WhatsApp / App CGE 1Click / Call Center / Correo Electrónico / Página Web).]": f"{texto_canal}.",
                        "[(Ej: nuestra Oficina Comercial / WhatsApp / App CGE 1Click / Call Center / Correo Electrónico / Página Web / Portal SEC)]": texto_canal,
                        
                        # Firma
                        "[Nombre y apellido Gerente Comercial]": GERENTES[zona],
                    }
                    
                    if tipo_carta == "error_lectura":
                        if rango_lectura:
                            # Cubrimos todas las variaciones de formato para los días
                            reemplazos["[XXXXXX y XXXXXX.]"] = f"{rango_lectura}."
                            reemplazos["[XXXXXX y XXXXXX]"] = rango_lectura
                            reemplazos["XXXXXX y XXXXXX"] = rango_lectura
                        
                        if fecha_boleta:
                            reemplazos["[día/mes/año]"] = fecha_boleta
                        if tipo_doc:
                            reemplazos["[boleta/factura]"] = tipo_doc
                        if consumo_kwh:
                            reemplazos["XXX kWh"] = f"{consumo_kwh} kWh"
                        if monto_boleta:
                            reemplazos["[$ XX.XXX]"] = monto_boleta
                    
                    # Ordenar por longitud
                    reemplazos_ordenados = sorted(reemplazos.items(), key=lambda x: len(x[0]), reverse=True)
                    
                    def aplicar_reemplazos(texto):
                        for key, value in reemplazos_ordenados:
                            texto = texto.replace(key, str(value))
                        return texto
                    
                    # Aplicar a párrafos
                    for paragraph in doc.paragraphs:
                        texto = paragraph.text
                        texto_nuevo = aplicar_reemplazos(texto)
                        
                        if texto_nuevo != texto:
                            paragraph.clear()
                            run = paragraph.add_run(texto_nuevo)
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
                            
# --- CONFIGURACIÓN DE NEGRITAS ---
                            # Definimos qué palabras SÍ deben ir en negrita (Destinatario y Referencia)
                            palabras_con_negrita = [
                                tratamiento, 
                                nombre_cliente, 
                                direccion, 
                                "Número de cliente:", 
                                "Ref.: Reclamo N°", 
                                GERENTES[zona], 
                                "COMPAÑÍA GENERAL DE ELECTRICIDAD S.A."
                            ]
                            
                            # Aplicamos negrita solo si el texto contiene elementos de la lista de arriba
                            # Quitamos 'comuna' y 'DGR N°' para que el encabezado superior sea normal
                            if any(x in texto_nuevo for x in palabras_con_negrita):
                                run.font.bold = True
                            else:
                                run.font.bold = False
                            
                            # Excepción para el cargo del gerente (que no sea negrita)
                            if "Gerente Comercial" in texto_nuevo and GERENTES[zona] not in texto_nuevo:
                                run.font.bold = False
                    
                    # Tablas
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    texto = aplicar_reemplazos(p.text)
                                    if texto != p.text:
                                        p.clear()
                                        r = p.add_run(texto)
                                        r.font.name = 'Arial'
                                        r.font.size = Pt(10)
                    
                    # Excel
                    if df is not None:
                        doc.add_page_break()
                        doc.add_heading("Anexo - Datos Adicionales", level=1)
                        table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
                        table.style = "Light Grid Accent 1"
                        for i, col in enumerate(df.columns):
                            cell = table.rows[0].cells[i]
                            cell.text = str(col)
                            for p in cell.paragraphs:
                                for r in p.runs:
                                    r.font.name = 'Arial'
                                    r.font.size = Pt(10)
                                    r.font.bold = True
                        for i, row in df.iterrows():
                            for j, value in enumerate(row):
                                cell = table.rows[i+1].cells[j]
                                cell.text = str(value)
                                for p in cell.paragraphs:
                                    for r in p.runs:
                                        r.font.name = 'Arial'
                                        r.font.size = Pt(10)
                    
                    # GUARDAR
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"Carta_{TIPOS_CARTAS[tipo_carta].replace(' ', '_')}_GR_{gr_numero}_{timestamp}.docx"
                    output_path = os.path.join("output", filename)
                    doc.save(output_path)
                    
                    # GUARDAMOS EN MEMORIA PARA QUE EL BOTÓN DE DESCARGA LO VEA
                    st.session_state.carta_generada = True
                    st.session_state.output_path = output_path
                    
                    # CREAMOS LA VISTA PREVIA EN MEMORIA TAMBIÉN
                    st.session_state.vista_previa_html = f"""
                        <div style='font-family:Arial;font-size:10pt;background:white;padding:20px;border:1px solid #ddd;border-radius:8px;color:black'>
                            <div style='text-align:right;margin-bottom:20px'>
                                <p style='margin:0'><strong>{comuna}</strong>, {hoy.day} de {meses[hoy.month]} de {hoy.year}</p>
                                <p style='margin:0'><strong>DGR N° {gr_numero} /{hoy.year}</strong></p>
                            </div>
                            <p><strong>{tratamiento}</strong></p>
                            <p><strong>{nombre_cliente}</strong></p>
                            <p><strong>{direccion}</strong></p>
                            <p><strong>{comuna}</strong></p>
                            <p style='margin-top:10px'><strong>Número de cliente: {numero_cliente}</strong></p>
                            <p><strong>Ref.: Reclamo N° {gr_numero}</strong></p>
                            <p style='margin-top:15px'>{estimado} {primer_nombre},</p>
                            <p style='text-align:justify'>Junto con saludar, le confirmamos que hemos recibido su reclamo por {TIPOS_CARTAS[tipo_carta].lower()}...</p>
                            <div style='margin-top:40px'>
                                <p><strong>{GERENTES[zona]}</strong></p>
                                <p>Gerente Comercial</p>
                                <p><strong>COMPAÑÍA GENERAL DE ELECTRICIDAD S.A.</strong></p>
                            </div>
                        </div>
                    """
                    
                    st.success("✅ ¡Carta generada exitosamente!")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# MOSTRAR VISTA PREVIA SI EXISTE EN MEMORIA
if st.session_state.get('carta_generada') and 'vista_previa_html' in st.session_state:
    with st.expander("👁️ VER VISTA PREVIA", expanded=True):
        st.markdown(st.session_state.vista_previa_html, unsafe_allow_html=True)

# FOOTER
st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📁 Cartas", len([f for f in os.listdir('output') if f.endswith('.docx')]))
with col2:
    st.metric("📄 Templates", len([f for f in os.listdir('templates') if f.endswith('.docx')]))
with col3:
    st.metric("⏱️ Tiempo", "~60 seg")

st.markdown("""
<div style='text-align:center;padding:20px;color:#64748b;font-size:0.9em'>
<p style='margin:5px 0'>⚡ <strong>Automatizador de Cartas </strong></p>
<p style='margin:5px 0'>Desarrollado por <strong>CiberByte</strong> | Javier Ruiz Arismendi</p>
<p style='margin:5px 0'>© 2025 - Sistema de Generación Automática</p>
</div>
""", unsafe_allow_html=True)
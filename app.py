import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os
from io import StringIO

# ================= CONFIG =================
st.set_page_config(
    page_title="Generador Cartas CGE",
    page_icon="📝",
    layout="wide"
)

os.makedirs("templates", exist_ok=True)
os.makedirs("output", exist_ok=True)

# ================= GERENTES =================
GERENTES = {
    "Norte": "Christian Alberto Gómez Díaz",
    "Centro": "Alex Andrés González Villablanca",
    "Sur": "Christian Enrique Araya Silva"
}

# ================= HEADER =================
st.markdown("""
<div style='background:linear-gradient(90deg,#1e40af,#0ea5e9);
padding:20px;border-radius:10px;margin-bottom:20px'>
<h1 style='color:white;margin:0'>📝 GENERADOR CARTA CGE</h1>
<p style='color:#e0f2fe;margin:5px 0 0 0'>Error de Lectura</p>
</div>
""", unsafe_allow_html=True)

# ================= FORMULARIO =================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📋 DATOS CLIENTE")

    ciudad_emision = st.text_input("Ciudad emisión:", "Valparaiso")
    comuna = st.text_input("Comuna:")
    tratamiento = st.selectbox("Tratamiento:", ["Señor", "Señora"])

    nombre_cliente = st.text_input("Nombre y apellido:")
    direccion = st.text_input("Dirección:")
    numero_cliente = st.text_input("Número cliente:")

    gr_numero = st.text_input("GR N°:")
    dgr_numero = st.text_input("DGR N°:")

    zona = st.selectbox("Zona:", ["Norte", "Centro", "Sur"])
    gerente = GERENTES[zona]

with col2:
    st.markdown("### 📊 DATOS FACTURACIÓN")

    canal_ingreso = st.selectbox(
        "Canal ingreso:",
        [
            "Oficina Comercial",
            "WhatsApp",
            "App CGE 1Click",
            "Call Center",
            "Correo Electrónico",
            "Página Web",
            "Portal SEC"
        ]
    )

    fecha_boleta = st.text_input("Fecha nueva boleta (dd/mm/aaaa):")
    tipo_doc = st.selectbox("Documento:", ["boleta", "factura"])

    consumo_kwh = st.text_input("Consumo corregido (kWh):")
    monto_boleta = st.text_input("Monto ($):")
    rango_lectura = st.text_input("Rango días lectura:", "06 y 12")

# ================= ANEXOS =================
st.markdown("### 📎 ANEXOS (opcional)")
tabla_excel = st.text_area("Pegar tabla Excel (TAB):", height=180)

df = None
if tabla_excel:
    try:
        df = pd.read_csv(StringIO(tabla_excel), sep="\t")
        st.dataframe(df, use_container_width=True)
    except:
        st.error("Formato incorrecto (usar TAB)")

# ================= GENERAR =================
if st.button("📝 GENERAR CARTA", type="primary"):

    if not nombre_cliente or not gr_numero:
        st.error("Faltan datos obligatorios")
    else:
        template_path = "templates/error_lectura.docx"

        if not os.path.exists(template_path):
            st.error("No existe el template error_lectura.docx")
        else:
            doc = Document(template_path)

            hoy = datetime.now()
            meses = [
                "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
            ]

            reemplazos = {
                "[Comuna]": ciudad_emision,
                "[día]": str(hoy.day),
                "[mes]": meses[hoy.month],
                "[202X]": str(hoy.year),
                "XXXXXXX": dgr_numero,

                "[Señor(a)]": tratamiento,
                "[Nombre y apellido reclamante]": nombre_cliente,
                "[Dirección]": direccion,
                "Número de cliente: 15965848":
                    f"Número de cliente: {numero_cliente}",

                "Ref.: Reclamo N° 15965848":
                    f"Ref.: Reclamo N° {gr_numero}",

                "[Estimado(a) Nombre,]":
                    f"Estimado(a) {nombre_cliente.split()[0]},",

                "[(Ej: nuestra Oficina Comercial / WhatsApp / App CGE 1Click / Call Center / Correo Electrónico / Página Web).]":
                    f"nuestro {canal_ingreso}.",

                "[boleta/factura]": tipo_doc,
                "[día/mes/año]": fecha_boleta,
                "XXX kWh": f"{consumo_kwh} kWh",
                "[$ XX.XXX]": monto_boleta,

                "[XXXXXX y XXXXXX]": rango_lectura,
                "[Nombre y apellido Gerente Comercial]": gerente
            }

            for p in doc.paragraphs:
                for k, v in reemplazos.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for k, v in reemplazos.items():
                            if k in cell.text:
                                cell.text = cell.text.replace(k, v)

            if df is not None:
                doc.add_page_break()
                doc.add_heading("Anexos", level=1)
                table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
                table.style = "Light Grid"
                for i, col in enumerate(df.columns):
                    table.rows[0].cells[i].text = col
                for i, row in df.iterrows():
                    for j, val in enumerate(row):
                        table.rows[i+1].cells[j].text = str(val)

            filename = f"Carta_Error_Lectura_GR_{gr_numero}.docx"
            path = os.path.join("output", filename)
            doc.save(path)

            st.success("✅ Carta generada correctamente")
            with open(path, "rb") as f:
                st.download_button(
                    "📥 Descargar Word editable",
                    f,
                    filename,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

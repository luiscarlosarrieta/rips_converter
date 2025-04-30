import streamlit as st
import pandas as pd
import json
from io import BytesIO


def safe_int(value):
    if pd.isna(value) or value == '':
        return None
    try:
        return int(value)
    except (ValueError, TypeError):
        return None

def safe_str(value):
    if pd.isna(value):
        return None
    return str(value)

def safe_float(value):
    if pd.isna(value):
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None

st.set_page_config(page_title="Convertidor Excel a JSON", layout="centered")
st.title("🧾 Convertidor Excel a JSON de Facturas")


def crear_json_desde_excel(file):
    # Cargar todas las hojas
    xls = pd.ExcelFile(file)
    transaccion_df = pd.read_excel(xls, sheet_name="transaccion")
    usuarios_df = pd.read_excel(xls, sheet_name="usuarios")
    consultas_df = pd.read_excel(xls, sheet_name="consultas")
    procedimientos_df = pd.read_excel(xls, sheet_name="procedimientos")

    transaccion_data = transaccion_df.iloc[0]

    resultado = {
        "numDocumentoIdObligado": str(transaccion_data["numDocumentoIdObligado"]),
        "numFactura": str(transaccion_data["numFactura"]),
        "tipoNota": str(transaccion_data["tipoNota"]),
        "numNota": str(transaccion_data["numNota"]),
        "usuarios": []
    }

    # Procesar usuarios
    for idx, usuario_row in usuarios_df.iterrows():
        num_doc = str(usuario_row["numDocumentoIdentificacion"])

        usuario = {
            "tipoDocumentoIdentificacion": usuario_row["tipoDocumentoIdentificacion"],
            "numDocumentoIdentificacion": num_doc,
            "consecutivo": idx + 1,
            "tipoUsuario": usuario_row["tipoUsuario"],
            "fechaNacimiento": str(usuario_row["fechaNacimiento"]),
            "codSexo": usuario_row["codSexo"],
            "codPaisResidencia": str(usuario_row["codPaisResidencia"]),
            "codMunicipioResidencia": str(usuario_row["codMunicipioResidencia"]),
            "codZonaTerritorialResidencia": str(usuario_row["codZonaTerritorialResidencia"]),
            "incapacidad": usuario_row["incapacidad"],
            "codPaisOrigen": str(usuario_row["codPaisOrigen"]),
            "servicios": {},
        }

        user_consultas = consultas_df[consultas_df["numDocumentoIdentificacion"].astype(str) == num_doc]
        if not user_consultas.empty:
            usuario["servicios"]["consultas"] = []
            for i, consulta_row in user_consultas.iterrows():
                consulta = {
                    "codPrestador": str(consulta_row["codPrestador"]),
                    "fechaInicioAtencion": str(consulta_row["fechaInicioAtencion"]),
                    "codConsulta": str(consulta_row["codConsulta"]),
                    "modalidadGrupoServicioTecSal": str(consulta_row["modalidadGrupoServicioTecSal"]),
                    "codServicio": safe_int(consulta_row["codServicio"]),
                    "grupoServicios": str(consulta_row["grupoServicios"]),
                    "finalidadTecnologiaSalud": str(consulta_row["finalidadTecnologiaSalud"]),
                    "causaMotivoAtencion": str(consulta_row["causaMotivoAtencion"]),
                    "tipoDiagnosticoPrincipal": str(consulta_row["tipoDiagnosticoPrincipal"]),
                    "tipoDocumentoIdentificacion": str(consulta_row["tipoDocumentoIdentificacion"]),
                    "conceptoRecaudo": str(consulta_row["conceptoRecaudo"]),
                    "vrServicio": int(float(consulta_row["vrServicio"])),
                    "numDocumentoIdentificacion": str(consulta_row["numDocumentoIdentificacion"]),
                    "valorPagoModerador": int(float(consulta_row["valorPagoModerador"])),
                    "numFEVPagoModerador": str(consulta_row["numFEVPagoModerador"]) if pd.notnull(consulta_row["numFEVPagoModerador"]) else "",
                    "consecutivo": i + 1,
                    "codDiagnosticoPrincipal": str(consulta_row["codDiagnosticoPrincipal"]),
                    "codDiagnosticoRelacionado1": str(consulta_row["codDiagnosticoRelacionado1"])
                }
                usuario["servicios"]["consultas"].append(consulta)

        user_procedimientos = procedimientos_df[procedimientos_df["numDocumentoIdentificacion"].astype(str) == num_doc]
        if not user_procedimientos.empty:
            usuario["procedimientos"] = []
            for j, procedimiento_row in user_procedimientos.iterrows():
                procedimiento = {
                    "codPrestador": str(procedimiento_row["codPrestador"]),
                    "fechaInicioAtencion": str(procedimiento_row["fechaInicioAtencion"]),
                    "idMIPRES": str(procedimiento_row["idMIPRES"]) if pd.notnull(procedimiento_row["idMIPRES"]) else None,
                    "numAutorizacion": str(procedimiento_row["numAutorizacion"]) if pd.notnull(procedimiento_row["numAutorizacion"]) else None,
                    "codProcedimiento": str(procedimiento_row["codProcedimiento"]),
                    "viaIngresoServicioSalud": str(procedimiento_row["viaIngresoServicioSalud"]),
                    "modalidadGrupoServicioTecSal": str(procedimiento_row["modalidadGrupoServicioTecSal"]),
                    "grupoServicios": str(procedimiento_row["grupoServicios"]),
                    "codServicio": int(procedimiento_row["codServicio"]),
                    "finalidadTecnologiaSalud": str(procedimiento_row["finalidadTecnologiaSalud"]),
                    "tipoDocumentoIdentificacion": str(procedimiento_row["tipoDocumentoIdentificacion"]),
                    "numDocumentoIdentificacion": str(procedimiento_row["numDocumentoIdentificacion"]),
                    "codDiagnosticoPrincipal": str(procedimiento_row["codDiagnosticoPrincipal"]),
                    "codDiagnosticoRelacionado": str(procedimiento_row["codDiagnosticoRelacionado"]),
                    "codComplicacion": str(procedimiento_row["codComplicacion"]),
                    "vrProcedimiento": int(float(procedimiento_row["vrProcedimiento"])),
                    "tipoPagoModerador": str(procedimiento_row["tipoPagoModerador"]),
                    "valorPagoModerador": int(float(procedimiento_row["valorPagoModerador"])),
                    "numFEVPagoModerador": str(procedimiento_row["numFEVPagoModerador"]) if pd.notnull(procedimiento_row["numFEVPagoModerador"]) else None,
                    "consecutivo": j + 1
                }
                usuario["procedimientos"].append(procedimiento)

        resultado["usuarios"].append(usuario)

    return resultado

st.subheader("📄 Descargar plantilla Excel")
columnas_transaccion = [
    "numDocumentoIdObligado", "numFactura", "tipoNota", "numNota"
]

columnas_usuarios = [
    "tipoDocumentoIdentificacion", "numDocumentoIdentificacion", "tipoUsuario",
    "fechaNacimiento", "codSexo", "codPaisResidencia", "codMunicipioResidencia",
    "codZonaTerritorialResidencia", "incapacidad", "codPaisOrigen"
]

columnas_consultas = [
    "codPrestador", "fechaInicioAtencion", "codConsulta", "modalidadGrupoServicioTecSal",
    "codServicio", "grupoServicios", "finalidadTecnologiaSalud", "causaMotivoAtencion",
    "tipoDiagnosticoPrincipal", "tipoDocumentoIdentificacion", "conceptoRecaudo",
    "vrServicio", "numDocumentoIdentificacion", "valorPagoModerador", "numFEVPagoModerador",
    "codDiagnosticoPrincipal", "codDiagnosticoRelacionado1"
]

columnas_procedimientos = [
    "codPrestador", "fechaInicioAtencion", "idMIPRES", "numAutorizacion",
    "codProcedimiento", "viaIngresoServicioSalud", "modalidadGrupoServicioTecSal",
    "grupoServicios", "codServicio", "finalidadTecnologiaSalud", "tipoDocumentoIdentificacion",
    "numDocumentoIdentificacion", "codDiagnosticoPrincipal", "codDiagnosticoRelacionado",
    "codComplicacion", "vrProcedimiento", "tipoPagoModerador", "valorPagoModerador",
    "numFEVPagoModerador"
]

transaccion_df = pd.DataFrame(columns=columnas_transaccion)
usuarios_df = pd.DataFrame(columns=columnas_usuarios)
consultas_df = pd.DataFrame(columns=columnas_consultas)
procedimientos_df = pd.DataFrame(columns=columnas_procedimientos)

# Creamos el buffer de memoria
buffer_plantilla = BytesIO()

# Escribimos las hojas al archivo Excel
with pd.ExcelWriter(buffer_plantilla, engine="openpyxl") as writer:
    transaccion_df.to_excel(writer, sheet_name="transaccion", index=False)
    usuarios_df.to_excel(writer, sheet_name="usuarios", index=False)
    consultas_df.to_excel(writer, sheet_name="consultas", index=False)
    procedimientos_df.to_excel(writer, sheet_name="procedimientos", index=False)

# Volvemos al inicio del buffer
buffer_plantilla.seek(0)
st.download_button(
    label="📥 Descargar plantilla Excel",
    data=buffer_plantilla,
    file_name="plantilla_factura.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("📤 Subir archivo Excel con datos")
uploaded_file = st.file_uploader("Selecciona un archivo Excel", type=["xlsx"])


if uploaded_file is not None:
    json_resultado = crear_json_desde_excel(uploaded_file)
    st.subheader("🧾 Vista previa del JSON generado")
    st.json(json_resultado)

    buffer_json = BytesIO()
    buffer_json.write(json.dumps(json_resultado, indent=2).encode())
    buffer_json.seek(0)

    st.download_button(
        label="📥 Descargar JSON",
        data=buffer_json,
        file_name="factura.json",
        mime="application/json"
    )
# if uploaded_file:
    # df = pd.read_excel(uploaded_file)
    # st.success("✅ Archivo cargado correctamente.")
    # json_resultado = crear_json_desde_excel(df)
    #
    # st.subheader("🧾 Vista previa del JSON generado")
    # st.json(json_resultado)
    #
    # buffer_json = BytesIO()
    # buffer_json.write(json.dumps(json_resultado, indent=2).encode())
    # buffer_json.seek(0)
    #
    # st.download_button(
    #     label="📥 Descargar JSON",
    #     data=buffer_json,
    #     file_name="factura.json",
    #     mime="application/json"
    # )

    # if not all(col in df.columns for col in columnas_plantilla):
    #     st.error("⚠️ El archivo no tiene todas las columnas requeridas.")
    # else:
    #     st.success("✅ Archivo cargado correctamente.")
    #     json_resultado = crear_json_desde_excel(df)
    #
    #     st.subheader("🧾 Vista previa del JSON generado")
    #     st.json(json_resultado)
    #
    #     buffer_json = BytesIO()
    #     buffer_json.write(json.dumps(json_resultado, indent=2).encode())
    #     buffer_json.seek(0)
    #
    #     st.download_button(
    #         label="📥 Descargar JSON",
    #         data=buffer_json,
    #         file_name="factura.json",
    #         mime="application/json"
    #     )

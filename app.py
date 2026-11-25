import streamlit as st
import pandas as pd
import json
import io
import smtplib
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Generación de PDFs idénticos a las planillas oficiales de Transnea
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Conexión a Supabase
try:
    from supabase import create_client
    SUPABASE_OK = True
except ImportError:
    SUPABASE_OK = False

# =============================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y CATÁLOGOS DE PERSONAL / VEHÍCULOS
# =============================================================================
st.set_page_config(
    page_title="Gestión MAPRO, Comisiones y Rendiciones | TRANSNEA",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
    .header-box {
        background: linear-gradient(135deg, #064E3B 0%, #059669 100%);
        padding: 18px 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .header-box h1 { margin: 0; font-size: 22px; font-weight: 800; }
    .header-box p { margin: 4px 0 0 0; font-size: 13px; color: #D1FAE5; }
</style>
""", unsafe_allow_html=True)

PERSONAL_MANTENIMIENTO = [
    "ESCOBAR HUGO", "SANDOVAL KEVIN", "FERNANDEZ CESAR", "GUTIERREZ MAURO",
    "FLOREZ RICARDO", "HANKE ERWIN", "RIVERO ARIEL", "ALARCON DARIO",
    "GONZALEZ GABRIEL", "VALLEJOS EDUARDO", "MAIDANA LARES THOMAS",
    "VARGAS LAUTARO", "CANDUSSO MATIAS", "CUSSIGH FABIAN", "OCAMPO JORGE",
    "BURGOS CARLOS", "MAIDANA WALTER", "ESQUIVEL MATIAS EXEQUIEL",
    "ORELLANA OSCAR IGNACIO", "ROMERO JUAN", "OLGUIN LUIS"
]

JEFES_CONSIGNA = ["MAIDANA WALTER", "CUSSIGH FABIAN", "LUIS AYALA (GFSA)"]

CUADRILLAS_BASE = ["BARRANQUERAS", "BELLA VISTA", "RESISTENCIA", "GOYA", "FORMOSA"]

VEHICULOS_FLOTA = [
    "-", "AI 208 SU", "AI 208 ST", "AH 973 WY", "AD 243 KQ",
    "AF 112 RT", "AE 890 LP", "AG 451 MN"
]

LISTA_LAT = [
    "SCA-BVI", "BVI-GOY", "RES-LLE", "LLE-MAN", "ECO-SMA",
    "RES-SCA", "RES-CCE", "RES-BAR1", "RES-BAR2", "GFO-PIR", "ET TUC"
]

DESTINATARIOS_MAPRO_DEFAULT = (
    "Jorge Luis Kochol; Gustavo Daniel Rovira; Pablo Gabriel Paredes; "
    "Guillermo Daniel Lemoine; Gabriel Nicolás Arias; Rodrigo Martin Garrido; "
    "Walter Andres Maidana; Matias Exequiel Esquivel; Cristian Adrian Vazquez"
)

DESTINATARIOS_CONTABLE_DEFAULT = (
    "Jorge Luis Kochol; Rodrigo Martin Garrido; Norma Beatriz Aguirre; "
    "Jainen Gabriel Encinas; Matias Exequiel Esquivel"
)

NUMEROS_LETRAS = {
    1: "UNO", 2: "DOS", 3: "TRES", 4: "CUATRO",
    5: "CINCO", 6: "SEIS", 7: "SIETE", 8: "OCHO"
}

DIAS_SEMANA_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# =============================================================================
# 2. CONEXIÓN A SUPABASE Y FUNCIONES CRUD
# =============================================================================
@st.cache_resource
def get_supabase():
    if SUPABASE_OK and "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    return None

sb = get_supabase()

def get_config():
    if sb:
        res = sb.table("config_viaticos").select("*").eq("id", 1).execute()
        if res.data:
            return res.data[0]
    return {
        "valor_viatico_diario": 55210.00,
        "ultimo_nro_comision": 154,
        "ultimo_nro_rendicion": 100,
        "solicitante_default": "Orellana Oscar Ignacio",
        "autorizante_default": "Ing. Gustavo Adolfo Paredes"
    }

def update_config(campos: dict):
    if sb:
        sb.table("config_viaticos").update(campos).eq("id", 1).execute()

def lunes_de_semana_iso(anio: int, semana: int) -> date:
    return date.fromisocalendar(anio, semana, 1)

# =============================================================================
# 3. PARSER DE EXCEL MAPRO SEMANAL (LAT - Zona Centro - SEM XX.xlsx)
# =============================================================================
def parsear_excel_mapro(file_bytes, semana_fallback, anio_fallback):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["LAT"] if "LAT" in wb.sheetnames else wb.active

    sem_val = ws["D4"].value or semana_fallback
    anio_val = ws["E4"].value or anio_fallback
    try:
        sem_int = int(sem_val)
    except Exception:
        sem_int = int(semana_fallback)
    try:
        anio_int = int(anio_val)
    except Exception:
        anio_int = int(anio_fallback)

    registros = []
    # Leer tabla principal (columnas B a R, filas 9 a 50)
    fecha_actual = None
    for r in range(9, 55):
        val_b = ws.cell(r, 2).value  # Día o Fecha
        lat_v = ws.cell(r, 3).value  # LAT
        if isinstance(val_b, datetime):
            fecha_actual = val_b.date()
        elif isinstance(val_b, date):
            fecha_actual = val_b

        if lat_v and str(lat_v).strip() not in ["", "0", "00:00:00"]:
            tct_si = str(ws.cell(r, 4).value or "").strip().upper()
            tct_flag = "SI" if tct_si == "X" else "NO"
            eq_cammesa = str(ws.cell(r, 6).value or "").strip()
            h_desde = str(ws.cell(r, 8).value or "07:00")[:5]
            h_hasta = str(ws.cell(r, 9).value or "14:00")[:5]
            t_rep = str(ws.cell(r, 10).value or "1 hora").strip()
            j_cons = str(ws.cell(r, 11).value or "MAIDANA WALTER").strip().upper()
            j_trab = str(ws.cell(r, 12).value or "").strip().upper()
            tipo_mto = str(ws.cell(r, 13).value or "Preventivo Semanal/Diario").strip()
            ot_num = str(ws.cell(r, 14).value or "").strip()
            tarea = str(ws.cell(r, 18).value or "").strip()

            f_usar = fecha_actual or lunes_de_semana_iso(anio_int, sem_int)
            dia_nombre = DIAS_SEMANA_ES[f_usar.weekday()]

            registros.append({
                "id_semana": f"{sem_int}-{anio_int}",
                "semana": sem_int,
                "anio": anio_int,
                "fecha": str(f_usar),
                "dia": dia_nombre,
                "lat": str(lat_v).strip().upper(),
                "tct": tct_flag,
                "equipo_cammesa": eq_cammesa,
                "horario_desde": h_desde,
                "horario_hasta": h_hasta,
                "tiempo_reposicion": t_rep,
                "jefe_consigna": j_cons,
                "jefe_trabajo": j_trab,
                "tipo_mto": tipo_mto,
                "tarea": tarea,
                "ot_num": ot_num,
                "requiere_comision": True,
                "comision_generada": False
            })

    return sem_int, anio_int, registros

# =============================================================================
# 4. GENERADORES DE PDF OFICIALES (COMISIÓN DE SERVICIO Y RENDICIÓN)
# =============================================================================
def generar_pdf_comision(com: dict, solicitante: str, autorizante: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=28, rightMargin=28, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    st_title = ParagraphStyle("T1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13, textColor=colors.HexColor("#064E3B"))
    st_sub = ParagraphStyle("T2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.HexColor("#1E293B"))
    st_small = ParagraphStyle("T3", parent=styles["Normal"], fontName="Helvetica", fontSize=8, textColor=colors.black)

    # Encabezado oficial
    story.append(Paragraph("TRANSNEA S.A. — COMISIONES DE SERVICIO", st_title))
    story.append(Paragraph("PERSONAL MANTENIMIENTO DE LÍNEAS", st_sub))
    story.append(Spacer(1, 8))

    f_emision = datetime.strptime(str(com["fecha_emision"])[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    f_ini_sem = datetime.strptime(str(com["inicio_semana"])[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    cant_ag = len(com["agentes_json"])
    cant_letras = NUMEROS_LETRAS.get(cant_ag, str(cant_ag))

    # Tabla superior de metadatos (igual al Excel Comision ORELLANA.xlsx)
    meta_data = [
        ["SEMANA:", str(com["semana"]), "NÚMERO DE COMISIÓN:", f"{com['nro_comision']} / {com['anio']}", "FECHA EMISIÓN:", f_emision],
        ["LAT / E.T.:", com["lat"], "Cuadrilla Base:", com["cuadrilla_base"], "Cant. Días:", f"{float(com['cant_dias']):.1f}".replace(".", ",")],
        ["O.T. Núm:", com.get("ot_num") or "-", "Vehículo 1:", com["vehiculo_1"], "1 VIÁTICO ($):", f"$ {float(com['valor_viatico']):,.2f}"],
        ["Jefe de Trab:", com["jefe_trabajo"], "Vehículo 2:", com["vehiculo_2"], "Gastos Extra ($):", f"$ {float(com['gastos_extra']):,.2f}"],
        ["Jefe de Cons:", com["jefe_consigna"], "Cant. Agentes:", f"{cant_ag} ({cant_letras})", "TOTAL ($):", f"$ {float(com['total_comision']):,.2f}"],
        ["TAREA A REALIZAR:", Paragraph(f"<b>{com['tarea']}</b>", st_small), "", "", "INICIO SEMANA:", f_ini_sem],
    ]

    t_meta = Table(meta_data, colWidths=[85, 95, 95, 95, 85, 85])
    t_meta.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#ECFDF5")),
        ("BACKGROUND", (2, 0), (2, 4), colors.HexColor("#ECFDF5")),
        ("BACKGROUND", (4, 0), (4, -1), colors.HexColor("#ECFDF5")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, 4), "Helvetica-Bold"),
        ("FONTNAME", (4, 0), (4, -1), "Helvetica-Bold"),
        ("FONTNAME", (5, 4), (5, 4), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("SPAN", (1, 5), (3, 5)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    # Nómina de agentes intervinientes
    story.append(Paragraph("NÓMINA DE AGENTES INTERVINIENTES", st_sub))
    story.append(Spacer(1, 4))

    dt_ini = datetime.strptime(str(com["inicio_semana"])[:10], "%Y-%m-%d").date()
    dias_headers = [
        f"Lunes {dt_ini.day}",
        f"Martes {(dt_ini + timedelta(days=1)).day}",
        f"Miércoles {(dt_ini + timedelta(days=2)).day}",
        f"Jueves {(dt_ini + timedelta(days=3)).day}",
        f"Viernes {(dt_ini + timedelta(days=4)).day}",
    ]

    nomina_rows = [["APELLIDO Y NOMBRE", "FUNCIÓN"] + dias_headers + ["TOTAL DÍAS", "IMPORTE ($)"]]
    viatico_unit = float(com["valor_viatico"])

    for ag in com["agentes_json"]:
        d_list = ag.get("dias_detalle", [0, 1, 1, 1, 0.5])
        tot_d = sum(d_list)
        imp_ag = tot_d * viatico_unit
        nomina_rows.append([
            ag["nombre"],
            ag.get("funcion", "Oficial"),
            str(d_list[0]).replace(".", ",") if d_list[0] else "",
            str(d_list[1]).replace(".", ",") if d_list[1] else "",
            str(d_list[2]).replace(".", ",") if d_list[2] else "",
            str(d_list[3]).replace(".", ",") if d_list[3] else "",
            str(d_list[4]).replace(".", ",") if d_list[4] else "",
            f"{tot_d:.1f}".replace(".", ","),
            f"$ {imp_ag:,.2f}"
        ])

    t_nom = Table(nomina_rows, colWidths=[145, 55, 45, 48, 52, 48, 48, 52, 67])
    t_nom.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#064E3B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (8, 1), (8, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_nom)
    story.append(Spacer(1, 16))

    # Firmas Solicitante y Autorización
    firmas_top = Table([
        ["SOLICITANTE", "", "AUTORIZACIÓN"],
        [f"\n\n....................................................\n{solicitante}", "", f"\n\n....................................................\n{autorizante}"]
    ], colWidths=[220, 100, 220])
    firmas_top.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
    ]))
    story.append(firmas_top)
    story.append(Spacer(1, 14))

    # Talones de Recibo de Viáticos por cada agente (2 columnas)
    story.append(Paragraph("RECIBOS INDIVIDUALES DE VIÁTICOS", st_sub))
    story.append(Spacer(1, 4))

    recibos_cells = []
    fila_actual = []
    for idx, ag in enumerate(com["agentes_json"]):
        tot_d = sum(ag.get("dias_detalle", [0, 1, 1, 1, 0.5]))
        imp_ag = tot_d * viatico_unit
        texto_recibo = (
            f"<b>RECIBÍ DE TRANSNEA S.A.</b> la suma de <b>$ {imp_ag:,.2f}</b> en concepto de "
            f"anticipo de viáticos ({tot_d:.1f} días) — Comisión N° {com['nro_comision']}/{com['anio']} ({com['lat']}).<br/><br/>"
            f"Firma: ........................................   Aclaración: <b>{ag['nombre']}</b>"
        )
        fila_actual.append(Paragraph(texto_recibo, st_small))
        if len(fila_actual) == 2:
            recibos_cells.append(fila_actual)
            fila_actual = []
    if fila_actual:
        fila_actual.append("")
        recibos_cells.append(fila_actual)

    t_rec = Table(recibos_cells, colWidths=[270, 270])
    t_rec.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t_rec)
    story.append(Spacer(1, 10))

    # Pie Total Gastos y Observaciones
    pie_data = [
        ["TOTAL GASTOS DE COMISIÓN ($):", f"$ {float(com['total_comision']):,.2f}"],
        ["OBSERVACIONES:", com.get("observaciones") or "-"]
    ]
    t_pie = Table(pie_data, colWidths=[200, 340])
    t_pie.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#ECFDF5")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t_pie)

    doc.build(story)
    return buf.getvalue()


def generar_pdf_rendicion(ren: dict, autorizante: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=28, rightMargin=28, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    st_title = ParagraphStyle("RT1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13, textColor=colors.HexColor("#064E3B"))
    st_sub = ParagraphStyle("RT2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.HexColor("#1E293B"))
    st_small = ParagraphStyle("RT3", parent=styles["Normal"], fontName="Helvetica", fontSize=8)

    f_ren = datetime.strptime(str(ren["fecha_rendicion"])[:10], "%Y-%m-%d").strftime("%d/%m/%Y")

    story.append(Paragraph(f"TRANSNEA S.A. — LIQUIDACIÓN DE GASTOS DE VIAJE N° {ren['nro_rendicion']} (Comisión N° {ren['nro_comision']}/{ren['anio']})", st_title))
    story.append(Spacer(1, 6))

    companeros_str = ", ".join(ren["companeros_json"]) if ren["companeros_json"] else "Sin acompañantes"
    cab_data = [
        ["Del Sr.:", ren["jefe_trabajo"], "LOCALIDAD:", ren["localidad_base"], "FECHA:", f_ren],
        ["Comisión realizada con:", Paragraph(companeros_str, st_small), "", "", "Vehículo:", ren["vehiculo"]],
    ]
    t_cab = Table(cab_data, colWidths=[95, 145, 70, 75, 65, 90])
    t_cab.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#ECFDF5")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#ECFDF5")),
        ("BACKGROUND", (4, 0), (4, -1), colors.HexColor("#ECFDF5")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, 0), "Helvetica-Bold"),
        ("FONTNAME", (4, 0), (4, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("SPAN", (1, 1), (3, 1)),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_cab)
    story.append(Spacer(1, 10))

    # Desarrollo de la comisión (Salida / Llegada)
    story.append(Paragraph("DESARROLLO DE LA COMISIÓN (ITINERARIO REAL)", st_sub))
    story.append(Spacer(1, 4))

    tramos_rows = [
        ["SALIDA", "", "", "", "LLEGADA", "", "", ""],
        ["LOCALIDAD", "DÍA", "MES", "HORA", "LOCALIDAD", "DÍA", "MES", "HORA"]
    ]
    for tr in ren["tramos_json"]:
        tramos_rows.append([
            tr["sal_loc"], str(tr["sal_dia"]), str(tr["sal_mes"]), str(tr["sal_hora"]),
            tr["lle_loc"], str(tr["lle_dia"]), str(tr["lle_mes"]), str(tr["lle_hora"])
        ])

    t_tr = Table(tramos_rows, colWidths=[130, 42, 42, 56, 130, 42, 42, 56])
    t_tr.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("SPAN", (0, 0), (3, 0)),
        ("SPAN", (4, 0), (7, 0)),
        ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#064E3B")),
        ("TEXTCOLOR", (0, 0), (-1, 1), colors.white),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_tr)
    story.append(Spacer(1, 12))

    # Detalle de gastos de viaje (Viáticos + Peajes + Otros comprobantes)
    story.append(Paragraph("DETALLE DE GASTOS DE VIAJE Y COMPROBANTES", st_sub))
    story.append(Spacer(1, 4))

    gastos_rows = [["FECHA", "Comp. Nro.", "CONCEPTO", "IMPORTE ($)"]]
    for g in ren["gastos_detalle_json"]:
        gastos_rows.append([
            str(g["fecha"]),
            str(g["comp_nro"]),
            str(g["concepto"]),
            f"$ {float(g['importe']):,.2f}"
        ])
    gastos_rows.append(["", "", "TOTAL RENDIDO ($)", f"$ {float(ren['total_rendido']):,.2f}"])

    t_gs = Table(gastos_rows, colWidths=[80, 95, 265, 100])
    t_gs.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#064E3B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#ECFDF5")),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (3, 0), (3, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_gs)
    story.append(Spacer(1, 8))

    dif = float(ren["diferencia_vs_comision"])
    estado_dif = "A favor del Agente (Reintegrar)" if dif > 0 else ("A devolver a Caja (Sobrante)" if dif < 0 else "Rendición Exacta ($ 0,00)")
    story.append(Paragraph(f"<b>Balance vs. Anticipo de Comisión:</b> $ {dif:,.2f} ({estado_dif})", st_small))
    story.append(Paragraph(f"<b>Observaciones / Tarea Cumplida:</b> {ren.get('observaciones') or '-'}", st_small))
    story.append(Spacer(1, 22))

    firmas_ren = Table([
        ["....................................................", "....................................................", "...................................................."],
        ["Firma de Jefe Inmediato", f"Autorizada por: {autorizante}", f"Firma Interesado: {ren['jefe_trabajo']}"]
    ], colWidths=[180, 180, 180])
    firmas_ren.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    story.append(firmas_ren)

    doc.build(story)
    return buf.getvalue()

# =============================================================================
# 5. ENVÍO DE CORREOS AUTOMÁTICOS VÍA SMTP (OUTLOOK / OFFICE 365)
# =============================================================================
def enviar_correo_smtp(destinatarios_str: str, cc_str: str, asunto: str, cuerpo_texto: str, adjuntos: list):
    if "SMTP_USER" not in st.secrets or "SMTP_PASSWORD" not in st.secrets:
        return False, "No hay credenciales SMTP configuradas en st.secrets."

    smtp_server = st.secrets.get("SMTP_SERVER", "smtp.office365.com")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    smtp_user = st.secrets["SMTP_USER"]
    smtp_pass = st.secrets["SMTP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = destinatarios_str
    if cc_str:
        msg["Cc"] = cc_str
    msg["Subject"] = asunto

    msg.attach(MIMEText(cuerpo_texto, "plain", "utf-8"))

    for nombre_arch, bytes_arch in adjuntos:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(bytes_arch)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{nombre_arch}"')
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=15)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        to_list = [x.strip() for x in destinatarios_str.replace(",", ";").split(";") if x.strip()]
        cc_list = [x.strip() for x in cc_str.replace(",", ";").split(";") if x.strip()] if cc_str else []
        server.sendmail(smtp_user, to_list + cc_list, msg.as_string())
        server.quit()
        return True, "Correo enviado exitosamente."
    except Exception as e:
        return False, f"Error SMTP: {e}"

# =============================================================================
# 6. INTERFAZ PRINCIPAL DE STREAMLIT
# =============================================================================
cfg = get_config()

st.markdown("""
<div class="header-box">
    <h1>⚡ TRANSNEA S.A. | Flujo Automatizado: MAPRO Semanal ➔ Comisiones ➔ Rendiciones</h1>
    <p>Programación de Mantenimiento LAT Zona Centro • Generación Automática de PDF Oficiales • Control de Viáticos y Peajes</p>
</div>
""", unsafe_allow_html=True)

if not sb:
    st.error("⚠️ Conectá tus credenciales `SUPABASE_URL` y `SUPABASE_KEY` en `.streamlit/secrets.toml` para operar con la base de datos en tiempo real.")
    st.stop()

# Selector global de Semana y Año
c_top1, c_top2, c_top3, c_top4 = st.columns([1.2, 1.2, 1.8, 2.2])
with c_top1:
    sem_sel = st.number_input("📅 Semana N°", min_value=1, max_value=53, value=49, step=1)
with c_top2:
    anio_sel = st.number_input("📆 Año", min_value=2025, max_value=2030, value=2026, step=1)
with c_top3:
    f_lunes = lunes_de_semana_iso(int(anio_sel), int(sem_sel))
    f_domingo = f_lunes + timedelta(days=6)
    st.info(f"Del **{f_lunes.strftime('%d/%m/%Y')}** al **{f_domingo.strftime('%d/%m/%Y')}**")
with c_top4:
    st.metric("💵 Valor Viático Diario Vigente", f"$ {float(cfg['valor_viatico_diario']):,.2f}", f"Últ. Comisión: N° {cfg['ultimo_nro_comision']}")

id_semana_actual = f"{int(sem_sel)}-{int(anio_sel)}"

tab_mapro, tab_comisiones, tab_rendiciones, tab_config = st.tabs([
    "1️⃣ Programación Semanal (MAPRO)",
    "2️⃣ Disparar Comisiones y PDFs",
    "3️⃣ Rendición de Gastos (Post-Tarea)",
    "⚙️ Tarifas de Viáticos y Config"
])

# =============================================================================
# PESTAÑA 1: PROGRAMACIÓN SEMANAL (MAPRO)
# =============================================================================
with tab_mapro:
    st.markdown(f"### 📋 Programación Semanal MAPRO — Semana {sem_sel} ({anio_sel})")

    col_imp, col_add = st.columns([1.2, 1.8])
    with col_imp:
        with st.container(border=True):
            st.markdown("#### 📥 Importar desde Excel MAPRO")
            st.caption("Subí tu planilla `LAT - Zona Centro - SEM XX.xlsx` para cargar las tareas automáticamente.")
            file_mapro = st.file_uploader("Seleccionar archivo MAPRO (.xlsx)", type=["xlsx"], key="upl_mapro")
            if file_mapro is not None:
                bytes_excel_mapro = file_mapro.read()
                s_imp, a_imp, regs_imp = parsear_excel_mapro(bytes_excel_mapro, sem_sel, anio_sel)
                st.success(f"Detectadas **{len(regs_imp)} filas** de programación para Semana **{s_imp}-{a_imp}**.")
                if st.button("🚀 Guardar MAPRO en Base de Datos", type="primary", use_container_width=True):
                    if regs_imp:
                        sb.table("programacion_semanal").insert(regs_imp).execute()
                        st.toast("¡MAPRO importada con éxito!", icon="✅")
                        st.rerun()

    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Agregar Tarea Manualmente a la Semana")
            ca1, ca2, ca3, ca4 = st.columns(4)
            with ca1:
                m_fecha = st.date_input("Fecha", value=f_lunes + timedelta(days=1), format="DD/MM/YYYY")
            with ca2:
                m_lat = st.selectbox("LAT / E.T.", LISTA_LAT)
            with ca3:
                m_jefe_t = st.selectbox("Jefe de Trabajo", PERSONAL_MANTENIMIENTO)
            with ca4:
                m_jefe_c = st.selectbox("Jefe de Consigna", JEFES_CONSIGNA)

            ca5, ca6, ca7, ca8 = st.columns([2.2, 1, 1, 1])
            with ca5:
                m_tarea = st.text_input("Tarea a realizar", placeholder="Ej: Limpieza de electroducto / Reparación de PAT")
            with ca6:
                m_tct = st.selectbox("TCT", ["NO", "SI"])
            with ca7:
                m_ot = st.text_input("O.T. N°", placeholder="Ej: OT-412")
            with ca8:
                m_req_com = st.checkbox("Lleva Comisión", value=True)

            if st.button("💾 Agregar a Programación", use_container_width=True):
                if not m_tarea.strip():
                    st.error("Ingresá la descripción de la tarea.")
                else:
                    sb.table("programacion_semanal").insert({
                        "id_semana": id_semana_actual,
                        "semana": int(sem_sel),
                        "anio": int(anio_sel),
                        "fecha": str(m_fecha),
                        "dia": DIAS_SEMANA_ES[m_fecha.weekday()],
                        "lat": m_lat,
                        "tct": m_tct,
                        "equipo_cammesa": "ANULACION DE RECIERRE" if m_tct == "SI" else "LICENCIA INTERNA",
                        "jefe_consigna": m_jefe_c,
                        "jefe_trabajo": m_jefe_t,
                        "tarea": m_tarea.strip(),
                        "ot_num": m_ot.strip(),
                        "requiere_comision": m_req_com,
                        "comision_generada": False
                    }).execute()
                    st.toast("Tarea agregada.", icon="✅")
                    st.rerun()

    # Mostrar programación cargada de la semana
    res_prog = sb.table("programacion_semanal").select("*").eq("id_semana", id_semana_actual).order("fecha").execute()
    df_prog = pd.DataFrame(res_prog.data) if res_prog.data else pd.DataFrame()

    if not df_prog.empty:
        st.markdown("#### 📋 Tareas Programadas en la Semana")
        st.dataframe(
            df_prog[["fecha", "dia", "lat", "tct", "jefe_consigna", "jefe_trabajo", "tarea", "ot_num", "requiere_comision", "comision_generada"]],
            use_container_width=True,
            hide_index=True
        )

        # Bloque de Envío de Correo MAPRO
        with st.expander("📧 Enviar Correo de Programación Semanal (MAPRO)", expanded=False):
            c_em1, c_em2 = st.columns(2)
            with c_em1:
                to_mapro = st.text_input("Para:", value=DESTINATARIOS_MAPRO_DEFAULT, key="to_mapro")
            with c_em2:
                cc_mapro = st.text_input("CC:", value="Oscar Ignacio Orellana", key="cc_mapro")

            asunto_mapro = st.text_input("Asunto:", value=f"Mapro semana {sem_sel} - LAT - Zona Centro", key="asunto_mapro")

            lineas_resumen = []
            for r in df_prog.itertuples():
                lineas_resumen.append(f"- {r.dia} ({r.fecha}): {r.lat} | {r.tarea} | JT: {r.jefe_trabajo} | JC: {r.jefe_consigna}")

            cuerpo_default_mapro = (
                f"Estimados, buenos días.\n\n"
                f"Envío adjunta la MAPRO correspondiente a la Semana {sem_sel} "
                f"(del {f_lunes.strftime('%d/%m/%Y')} al {f_domingo.strftime('%d/%m/%Y')}) de Mantenimiento de LAT Zona Centro.\n\n"
                f"Detalle de trabajos:\n" + "\n".join(lineas_resumen) + "\n\n"
                f"Quedamos a disposición.\nMuchas gracias.\nSaludos cordiales.\n\n"
                f"{cfg['solicitante_default']}\nGerencia de Mantenimiento - Mantenimiento de Líneas"
            )
            cuerpo_mapro = st.text_area("Cuerpo del correo:", value=cuerpo_default_mapro, height=200, key="body_mapro")

            adj_mapro = []
            if "bytes_excel_mapro" in locals():
                adj_mapro.append((f"LAT - Zona Centro - SEM {sem_sel}.xlsx", bytes_excel_mapro))

            if st.button("📨 Disparar Email MAPRO Ahora", type="primary", key="btn_send_mapro"):
                ok, msg = enviar_correo_smtp(to_mapro, cc_mapro, asunto_mapro, cuerpo_mapro, adj_mapro)
                if ok:
                    st.success(msg)
                else:
                    st.warning(f"{msg} (Podés copiar el texto de arriba y pegarlo en Outlook).")

# =============================================================================
# PESTAÑA 2: DISPARAR COMISIONES DE SERVICIO Y GENERAR PDFs
# =============================================================================
with tab_comisiones:
    st.markdown(f"### 🧾 Generación Automática de Comisiones — Semana {sem_sel} ({anio_sel})")
    st.caption("El sistema agrupa automáticamente las tareas de la MAPRO por Jefe de Trabajo y Línea para emitir la Comisión oficial en PDF.")

    res_pend = sb.table("programacion_semanal").select("*").eq("id_semana", id_semana_actual).eq("requiere_comision", True).execute()
    df_pend = pd.DataFrame(res_pend.data) if res_pend.data else pd.DataFrame()

    if df_pend.empty:
        st.info("No hay tareas cargadas con comisión pendiente para esta semana. Importá o cargá la MAPRO en la pestaña 1.")
    else:
        # Agrupar por Jefe de Trabajo y LAT
        grupos = df_pend.groupby(["jefe_trabajo", "lat"], as_index=False).agg({
            "tarea": lambda x: " / ".join(sorted(set(x))),
            "jefe_consigna": "first",
            "ot_num": "first",
            "fecha": "count"
        }).rename(columns={"fecha": "dias_detectados"})

        st.markdown("#### ⚡ Cuadrillas Detectadas en la Programación para Disparar Comisión")
        sel_idx = st.selectbox(
            "Seleccioná la cuadrilla / frente de trabajo a configurar:",
            range(len(grupos)),
            format_func=lambda i: f"JT: {grupos.loc[i, 'jefe_trabajo']}  |  LAT: {grupos.loc[i, 'lat']}  |  Tarea: {grupos.loc[i, 'tarea']}"
        )

        g_sel = grupos.loc[sel_idx]

        with st.form("form_generar_comision"):
            fc1, fc2, fc3, fc4 = st.columns(4)
            with fc1:
                nro_com_nuevo = st.number_input("N° de Comisión", min_value=1, value=int(cfg["ultimo_nro_comision"]) + 1, step=1)
            with fc2:
                cuad_base = st.selectbox("Cuadrilla Base", CUADRILLAS_BASE, index=1 if g_sel["lat"] == "BVI-GOY" else 0)
            with fc3:
                veh_1 = st.selectbox("Vehículo 1", VEHICULOS_FLOTA, index=1)
            with fc4:
                veh_2 = st.selectbox("Vehículo 2", VEHICULOS_FLOTA, index=0)

            fc5, fc6, fc7, fc8 = st.columns(4)
            with fc5:
                cant_dias_com = st.number_input("Cantidad de Días Viático", min_value=0.5, max_value=15.0, value=3.5, step=0.5)
            with fc6:
                val_viatico_com = st.number_input("Valor 1 Viático ($)", min_value=1000.0, value=float(cfg["valor_viatico_diario"]), step=500.0)
            with fc7:
                gastos_ext_com = st.number_input("Gastos Extra / Peajes ($)", min_value=0.0, value=0.0, step=500.0)
            with fc8:
                f_emision_com = st.date_input("Fecha de Emisión", value=date.today(), format="DD/MM/YYYY")

            tarea_edit = st.text_input("Tarea a Realizar (Editable)", value=str(g_sel["tarea"]).upper())
            obs_com = st.text_input("Observaciones", placeholder="Opcional")

            st.markdown("##### 👷 Nómina de Agentes Intervinientes (Seleccioná los integrantes de la cuadrilla)")
            agentes_default = [g_sel["jefe_trabajo"]] if g_sel["jefe_trabajo"] in PERSONAL_MANTENIMIENTO else []
            agentes_sel = st.multiselect(
                "Agentes de la cuadrilla:",
                options=PERSONAL_MANTENIMIENTO,
                default=agentes_default
            )

            submitted_com = st.form_submit_button("📄 Generar Comisión y Guardar en Base de Datos", type="primary", use_container_width=True)

            if submitted_com:
                if not agentes_sel:
                    st.error("Seleccioná al menos 1 agente en la nómina.")
                else:
                    # Construir detalle diario estándar según cant_dias_com
                    if cant_dias_com == 4.5:
                        d_det = [1, 1, 1, 1, 0.5]
                    elif cant_dias_com == 3.5:
                        d_det = [0, 1, 1, 1, 0.5]
                    else:
                        d_det = [round(cant_dias_com / 5, 2)] * 5

                    agentes_payload = [{"nombre": a, "funcion": "Oficial", "dias_detalle": d_det} for a in agentes_sel]
                    total_calc = (len(agentes_sel) * float(cant_dias_com) * float(val_viatico_com)) + float(gastos_ext_com)

                    nueva_com = {
                        "nro_comision": int(nro_com_nuevo),
                        "anio": int(anio_sel),
                        "id_semana": id_semana_actual,
                        "semana": int(sem_sel),
                        "fecha_emision": str(f_emision_com),
                        "inicio_semana": str(f_lunes),
                        "lat": g_sel["lat"],
                        "ot_num": str(g_sel["ot_num"] or ""),
                        "cuadrilla_base": cuad_base,
                        "vehiculo_1": veh_1,
                        "vehiculo_2": veh_2,
                        "jefe_trabajo": g_sel["jefe_trabajo"],
                        "jefe_consigna": g_sel["jefe_consigna"],
                        "tarea": tarea_edit,
                        "cant_dias": float(cant_dias_com),
                        "valor_viatico": float(val_viatico_com),
                        "gastos_extra": float(gastos_ext_com),
                        "total_comision": round(total_calc, 2),
                        "agentes_json": agentes_payload,
                        "observaciones": obs_com,
                        "estado": "Emitida",
                        "rendicion_realizada": False
                    }

                    sb.table("comisiones_servicio").insert(nueva_com).execute()
                    update_config({"ultimo_nro_comision": int(nro_com_nuevo)})
                    sb.table("programacion_semanal").update({"comision_generada": True}).eq("id_semana", id_semana_actual).eq("jefe_trabajo", g_sel["jefe_trabajo"]).eq("lat", g_sel["lat"]).execute()

                    st.toast(f"¡Comisión N° {nro_com_nuevo} generada correctamente!", icon="🎉")
                    st.rerun()

    st.divider()

    # Listado de Comisiones Emitidas en la Semana + Descarga PDF + Email Contable
    res_coms = sb.table("comisiones_servicio").select("*").eq("id_semana", id_semana_actual).order("nro_comision").execute()
    lista_coms = res_coms.data if res_coms.data else []

    if lista_coms:
        st.markdown(f"#### 📂 Comisiones Emitidas — Semana {sem_sel} ({len(lista_coms)} comisiones)")
        pdfs_para_mail = []
        bloques_texto_mail = []

        for c in lista_coms:
            pdf_bytes = generar_pdf_comision(c, cfg["solicitante_default"], cfg["autorizante_default"])
            apellido_jt = str(c["jefe_trabajo"]).split()[0]
            nombre_pdf = f"SEM {c['semana']} - {c['lat']} - {apellido_jt}.pdf"
            pdfs_para_mail.append((nombre_pdf, pdf_bytes))

            bloques_texto_mail.append(
                f"Comisión {c['nro_comision']}\n"
                f"- Jefe de Trabajo: {c['jefe_trabajo']}\n"
                f"- Tarea: {c['tarea']}\n"
                f"- LAT: {c['lat']}\n"
                f"- Agentes: {len(c['agentes_json'])}\n"
                f"- Días: {str(c['cant_dias']).replace('.', ',')}\n"
            )

            col_c1, col_c2, col_c3 = st.columns([2.5, 1.2, 1.3])
            with col_c1:
                st.write(f"**Comisión N° {c['nro_comision']}** | LAT: `{c['lat']}` | JT: **{c['jefe_trabajo']}** | Agentes: **{len(c['agentes_json'])}** | Total: **$ {float(c['total_comision']):,.2f}**")
            with col_c2:
                st.download_button(
                    label=f"⬇️ Descargar {nombre_pdf}",
                    data=pdf_bytes,
                    file_name=nombre_pdf,
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"dl_com_{c['id']}"
                )
            with col_c3:
                if st.button("🗑️ Eliminar", key=f"del_com_{c['id']}", use_container_width=True):
                    sb.table("comisiones_servicio").delete().eq("id", c["id"]).execute()
                    st.rerun()

        with st.expander("📧 Disparar Email de Pedidos de Viáticos al Área Contable", expanded=True):
            ce1, ce2 = st.columns(2)
            with ce1:
                to_cont = st.text_input("Para (Contable):", value=DESTINATARIOS_CONTABLE_DEFAULT, key="to_contable")
            with ce2:
                cc_cont = st.text_input("CC:", value="Oscar Ignacio Orellana", key="cc_contable")

            asunto_cont = st.text_input("Asunto:", value=f"Pedidos de viáticos - Semana {sem_sel}", key="asunto_contable")
            cuerpo_cont_default = (
                f"Estimados, buenos días.\n\n"
                f"Envío los pedidos de comisiones para la semana {sem_sel} de acuerdo con la programación:\n\n"
                + "\n".join(bloques_texto_mail) +
                f"\nQuedo a su disposición, saludos cordiales.\n\n"
                f"{cfg['solicitante_default']}\nGerencia de Mantenimiento - Mantenimiento de Líneas"
            )
            cuerpo_cont = st.text_area("Cuerpo del correo:", value=cuerpo_cont_default, height=240, key="body_contable")

            if st.button("📨 Enviar Correo con Todos los PDFs Adjuntos", type="primary", use_container_width=True):
                ok_c, msg_c = enviar_correo_smtp(to_cont, cc_cont, asunto_cont, cuerpo_cont, pdfs_para_mail)
                if ok_c:
                    st.success(msg_c)
                else:
                    st.warning(f"{msg_c} (Descargá los PDFs de arriba y pegá este texto en Outlook).")

# =============================================================================
# PESTAÑA 3: RENDICIÓN DE GASTOS DE VIAJE (SEMANA SIGUIENTE)
# =============================================================================
with tab_rendiciones:
    st.markdown("### 🧾 Liquidación de Gastos de Viaje (Rendición Post-Tarea)")
    st.caption("Seleccioná una comisión emitida para liquidar los días reales cumplidos, horarios de salida/llegada, peajes y comprobantes.")

    res_all_coms = sb.table("comisiones_servicio").select("*").order("nro_comision", desc=True).limit(40).execute()
    coms_disp = res_all_coms.data if res_all_coms.data else []

    if not coms_disp:
        st.info("Aún no hay comisiones emitidas para rendir.")
    else:
        com_sel_ren = st.selectbox(
            "Seleccionar Comisión a Rendir:",
            coms_disp,
            format_func=lambda c: f"Comisión N° {c['nro_comision']}/{c['anio']} (Sem {c['semana']}) | {c['lat']} | JT: {c['jefe_trabajo']} | Anticipo: $ {float(c['total_comision']):,.2f} {'✅ Rendida' if c['rendicion_realizada'] else '⏳ Pendiente'}"
        )

        with st.container(border=True):
            st.markdown(f"#### 📝 Formulario de Liquidación — Comisión N° {com_sel_ren['nro_comision']}/{com_sel_ren['anio']}")
            rc1, rc2, rc3, rc4 = st.columns(4)
            with rc1:
                nro_ren_nuevo = st.number_input("N° Liquidación / Rendición", min_value=1, value=int(cfg["ultimo_nro_rendicion"]) + 1, step=1)
            with rc2:
                loc_base_ren = st.text_input("Localidad Base", value="Bqueras.")
            with rc3:
                f_ren_fecha = st.date_input("Fecha Rendición", value=date.today(), format="DD/MM/YYYY")
            with rc4:
                veh_ren = st.text_input("Vehículo Utilizado", value=com_sel_ren["vehiculo_1"])

            # Tramos de Salida y Llegada (Ida y Vuelta)
            st.markdown("##### 🚐 Desarrollo de la Comisión (Itinerario Real de Salida y Llegada)")
            dt_ini_c = datetime.strptime(str(com_sel_ren["inicio_semana"])[:10], "%Y-%m-%d").date()
            dt_fin_c = dt_ini_c + timedelta(days=4)

            df_tramos_init = pd.DataFrame([
                {"sal_loc": "Barranqueras", "sal_dia": dt_ini_c.day, "sal_mes": dt_ini_c.month, "sal_hora": "07:00",
                 "lle_loc": com_sel_ren["lat"], "lle_dia": dt_ini_c.day, "lle_mes": dt_ini_c.month, "lle_hora": "10:00"},
                {"sal_loc": com_sel_ren["lat"], "sal_dia": dt_fin_c.day, "sal_mes": dt_fin_c.month, "sal_hora": "10:30",
                 "lle_loc": "Barranqueras", "lle_dia": dt_fin_c.day, "lle_mes": dt_fin_c.month, "lle_hora": "15:00"},
            ])
            df_tramos_ed = st.data_editor(df_tramos_init, num_rows="dynamic", use_container_width=True, key="ed_tramos_ren")

            # Detalle de gastos (Viáticos reales + Peajes + Combustible/Otros)
            st.markdown("##### 💵 Detalle de Gastos (Viáticos Reales + Peajes + Comprobantes)")
            cant_ag_ren = len(com_sel_ren["agentes_json"])
            df_gastos_init = pd.DataFrame([
                {
                    "fecha": dt_ini_c.strftime("%d/%m/%Y"),
                    "comp_nro": "S/C",
                    "concepto": f"${float(com_sel_ren['valor_viatico']):,.0f}/día x {com_sel_ren['cant_dias']} días x {cant_ag_ren} agentes",
                    "importe": float(com_sel_ren["total_comision"])
                },
                {
                    "fecha": dt_ini_c.strftime("%d/%m/%Y"),
                    "comp_nro": "",
                    "concepto": "Peaje puente Chaco-Corrientes (Ida)",
                    "importe": 1500.0
                },
                {
                    "fecha": dt_fin_c.strftime("%d/%m/%Y"),
                    "comp_nro": "",
                    "concepto": "Peaje puente Chaco-Corrientes (Vuelta)",
                    "importe": 1500.0
                }
            ])
            df_gastos_ed = st.data_editor(df_gastos_init, num_rows="dynamic", use_container_width=True, key="ed_gastos_ren")

            total_rendido_calc = float(pd.to_numeric(df_gastos_ed["importe"], errors="coerce").fillna(0).sum())
            dif_calc = total_rendido_calc - float(com_sel_ren["total_comision"])

            km1, km2, km3 = st.columns(3)
            km1.metric("Anticipo Entregado (Comisión)", f"$ {float(com_sel_ren['total_comision']):,.2f}")
            km2.metric("Total Rendido Real", f"$ {total_rendido_calc:,.2f}")
            km3.metric("Diferencia (Reintegro / Devolución)", f"$ {dif_calc:,.2f}")

            obs_ren = st.text_input("Observaciones de Rendición", value=f"{com_sel_ren['tarea']} — LAT {com_sel_ren['lat']}")

            if st.button("💾 Guardar Rendición y Generar PDF Oficial", type="primary", use_container_width=True):
                companeros = [a["nombre"] for a in com_sel_ren["agentes_json"] if a["nombre"] != com_sel_ren["jefe_trabajo"]]
                nueva_ren = {
                    "nro_rendicion": int(nro_ren_nuevo),
                    "comision_id": com_sel_ren["id"],
                    "nro_comision": com_sel_ren["nro_comision"],
                    "anio": com_sel_ren["anio"],
                    "fecha_rendicion": str(f_ren_fecha),
                    "localidad_base": loc_base_ren,
                    "jefe_trabajo": com_sel_ren["jefe_trabajo"],
                    "vehiculo": veh_ren,
                    "companeros_json": companeros,
                    "tramos_json": df_tramos_ed.to_dict(orient="records"),
                    "gastos_detalle_json": df_gastos_ed.to_dict(orient="records"),
                    "total_rendido": round(total_rendido_calc, 2),
                    "diferencia_vs_comision": round(dif_calc, 2),
                    "observaciones": obs_ren
                }
                sb.table("rendiciones_gastos").insert(nueva_ren).execute()
                sb.table("comisiones_servicio").update({"rendicion_realizada": True, "estado": "Rendida"}).eq("id", com_sel_ren["id"]).execute()
                update_config({"ultimo_nro_rendicion": int(nro_ren_nuevo)})

                st.toast(f"¡Rendición N° {nro_ren_nuevo} guardada!", icon="✅")
                st.rerun()

    # Historial de Rendiciones Guardadas
    res_rens = sb.table("rendiciones_gastos").select("*").order("nro_rendicion", desc=True).limit(20).execute()
    if res_rens.data:
        st.divider()
        st.markdown("#### 📑 Historial de Rendiciones Emitidas")
        for r in res_rens.data:
            pdf_ren_bytes = generar_pdf_rendicion(r, cfg["autorizante_default"])
            apellido_r = str(r["jefe_trabajo"]).split()[0]
            fname_r = f"RENDICION_{r['nro_rendicion']}_COM_{r['nro_comision']}_{apellido_r}.pdf"
            cr1, cr2 = st.columns([3.5, 1.5])
            with cr1:
                st.write(f"**Liquidación N° {r['nro_rendicion']}** (Comisión {r['nro_comision']}/{r['anio']}) | JT: **{r['jefe_trabajo']}** | Total Rendido: **$ {float(r['total_rendido']):,.2f}** | Dif: **$ {float(r['diferencia_vs_comision']):,.2f}**")
            with cr2:
                st.download_button(
                    label="⬇️ Descargar PDF Rendición",
                    data=pdf_ren_bytes,
                    file_name=fname_r,
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"dl_ren_{r['id']}"
                )

# =============================================================================
# PESTAÑA 4: CONFIGURACIÓN DE TARIFAS Y NUMERACIÓN CORRELATIVA
# =============================================================================
with tab_config:
    st.markdown("### ⚙️ Parámetros Globales de Viáticos y Numeración")
    st.caption("Cuando se actualice por paritarias el monto del viático diario, modificalo acá una sola vez y se aplicará a todas las comisiones nuevas.")
    with st.form("form_cfg_viaticos"):
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            nuevo_viatico = st.number_input("Valor Viático Diario ($)", min_value=1000.0, value=float(cfg["valor_viatico_diario"]), step=500.0)
        with cc2:
            ult_com = st.number_input("Último N° Comisión Emitido", min_value=1, value=int(cfg["ultimo_nro_comision"]), step=1)
        with cc3:
            ult_ren = st.number_input("Último N° Rendición Emitido", min_value=1, value=int(cfg["ultimo_nro_rendicion"]), step=1)

        cc4, cc5 = st.columns(2)
        with cc4:
            sol_def = st.text_input("Solicitante por defecto", value=cfg["solicitante_default"])
        with cc5:
            aut_def = st.text_input("Autorizante por defecto", value=cfg["autorizante_default"])

        if st.form_submit_button("💾 Guardar Configuración", type="primary"):
            update_config({
                "valor_viatico_diario": float(nuevo_viatico),
                "ultimo_nro_comision": int(ult_com),
                "ultimo_nro_rendicion": int(ult_ren),
                "solicitante_default": sol_def.strip(),
                "autorizante_default": aut_def.strip()
            })
            st.success("Configuración actualizada correctamente.")
            st.rerun()

import streamlit as st
import pandas as pd
import io
from datetime import date, datetime, timedelta
import openpyxl

# Generación de PDFs oficiales de Transnea
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
        return create_client(str(st.secrets["SUPABASE_URL"]).strip(), str(st.secrets["SUPABASE_KEY"]).strip())
    return None

sb = get_supabase()

def get_config():
    if sb:
        try:
            res = sb.table("config_viaticos").select("*").eq("id", 1).execute()
            if res.data:
                return res.data[0]
        except Exception as e:
            st.error(f"Error leyendo config_viaticos en Supabase: {e}")
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

    sem_val = ws["D4"].value
    try:
        sem_int = int(float(sem_val)) if sem_val is not None else int(semana_fallback)
    except Exception:
        sem_int = int(semana_fallback)

    anio_int = int(anio_fallback)

    registros = []
    fecha_actual = None
    for r in range(9, 55):
        val_b = ws.cell(r, 2).value
        lat_v = ws.cell(r, 3).value

        if isinstance(val_b, datetime):
            fecha_actual = val_b.date()
        elif isinstance(val_b, date):
            fecha_actual = val_b

        if lat_v is not None:
            lat_str = str(lat_v).strip().upper()
            if lat_str and lat_str not in ["0", "00:00:00", "NONE", "NAN"]:
                tct_si = str(ws.cell(r, 4).value or "").strip().upper()
                tct_flag = "SI" if tct_si == "X" else "NO"
                eq_cammesa = str(ws.cell(r, 6).value or "").strip()
                h_desde = str(ws.cell(r, 8).value or "07:00")[:5]
                h_hasta = str(ws.cell(r, 9).value or "14:00")[:5]
                t_rep = str(ws.cell(r, 10).value or "1 hora").strip()
                j_cons = str(ws.cell(r, 11).value or "MAIDANA WALTER").strip().upper()
                j_trab = str(ws.cell(r, 12).value or "SIN ASIGNAR").strip().upper()
                tipo_mto = str(ws.cell(r, 13).value or "Preventivo Semanal/Diario").strip()
                ot_raw = ws.cell(r, 14).value
                ot_num = "" if ot_raw is None else str(ot_raw).strip()
                tarea_raw = ws.cell(r, 18).value
                tarea = str(tarea_raw).strip() if tarea_raw is not None and str(tarea_raw).strip() else "Mantenimiento programado LAT"

                f_usar = fecha_actual if isinstance(fecha_actual, date) else lunes_de_semana_iso(anio_int, sem_int)
                dia_nombre = DIAS_SEMANA_ES[f_usar.weekday()]

                registros.append({
                    "id_semana": f"{sem_int}-{anio_int}",
                    "semana": int(sem_int),
                    "anio": int(anio_int),
                    "fecha": f_usar.isoformat(),
                    "dia": str(dia_nombre),
                    "lat": lat_str,
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
# 4. GENERADORES DE PDF SEMANALES (COMISIÓN DE SERVICIO Y RENDICIÓN)
# =============================================================================
def generar_pdf_comision(com: dict, solicitante: str, autorizante: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=28, rightMargin=28, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    st_title = ParagraphStyle("T1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13, textColor=colors.HexColor("#064E3B"))
    st_sub = ParagraphStyle("T2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.HexColor("#1E293B"))
    st_small = ParagraphStyle("T3", parent=styles["Normal"], fontName="Helvetica", fontSize=8, textColor=colors.black)

    f_emision = datetime.strptime(str(com["fecha_emision"])[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    f_ini_sem = datetime.strptime(str(com["inicio_semana"])[:10], "%Y-%m-%d").date()
    f_fin_sem = f_ini_sem + timedelta(days=6)
    rango_semanal = f"Sem {com['semana']} ({f_ini_sem.strftime('%d/%m/%Y')} al {f_fin_sem.strftime('%d/%m/%Y')})"

    cant_ag = len(com["agentes_json"])
    cant_letras = NUMEROS_LETRAS.get(cant_ag, str(cant_ag))
    dias_sem = float(com["cant_dias"])
    viatico_unit = float(com["valor_viatico"])

    story.append(Paragraph("TRANSNEA S.A. — COMISIÓN DE SERVICIO (RESUMEN SEMANAL)", st_title))
    story.append(Paragraph(f"PERSONAL MANTENIMIENTO DE LÍNEAS — {rango_semanal}", st_sub))
    story.append(Spacer(1, 8))

    meta_data = [
        ["SEMANA:", str(com["semana"]), "NÚMERO DE COMISIÓN:", f"{com['nro_comision']} / {com['anio']}", "FECHA EMISIÓN:", f_emision],
        ["LAT / E.T.:", com["lat"], "Cuadrilla Base:", com["cuadrilla_base"], "Días en Semana:", f"{dias_sem:.1f}".replace(".", ",")],
        ["O.T. Núm:", com.get("ot_num") or "-", "Vehículo 1:", com["vehiculo_1"], "Viático Unit. ($):", f"$ {viatico_unit:,.2f}"],
        ["Jefe de Trab:", com["jefe_trabajo"], "Vehículo 2:", com["vehiculo_2"], "Gastos Extra ($):", f"$ {float(com['gastos_extra']):,.2f}"],
        ["Jefe de Cons:", com["jefe_consigna"], "Cant. Agentes:", f"{cant_ag} ({cant_letras})", "TOTAL SEMANAL ($):", f"$ {float(com['total_comision']):,.2f}"],
        ["TAREA SEMANAL:", Paragraph(f"<b>{com['tarea']}</b>", st_small), "", "", "PERÍODO:", f"{f_ini_sem.strftime('%d/%m')} - {f_fin_sem.strftime('%d/%m/%Y')}"],
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

    # Nómina consolidada semanal (sin columnas día por día)
    story.append(Paragraph("NÓMINA DE AGENTES — LIQUIDACIÓN SEMANAL CONSOLIDADA", st_sub))
    story.append(Spacer(1, 4))

    nomina_rows = [["APELLIDO Y NOMBRE", "FUNCIÓN", "PERÍODO SEMANAL", "DÍAS SEMANA", "IMPORTE SEMANAL ($)"]]
    for ag in com["agentes_json"]:
        d_ag = float(ag.get("dias_semana", dias_sem))
        imp_ag = d_ag * viatico_unit
        nomina_rows.append([
            ag["nombre"],
            ag.get("funcion", "Oficial"),
            f"Sem {com['semana']} ({f_ini_sem.strftime('%d/%m')} al {f_fin_sem.strftime('%d/%m')})",
            f"{d_ag:.1f}".replace(".", ","),
            f"$ {imp_ag:,.2f}"
        ])

    t_nom = Table(nomina_rows, colWidths=[165, 70, 130, 75, 100])
    t_nom.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#064E3B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 1), (4, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_nom)
    story.append(Spacer(1, 16))

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

    # Recibos semanales por agente
    story.append(Paragraph("RECIBOS SEMANALES DE VIÁTICOS POR AGENTE", st_sub))
    story.append(Spacer(1, 4))

    recibos_cells = []
    fila_actual = []
    for ag in com["agentes_json"]:
        d_ag = float(ag.get("dias_semana", dias_sem))
        imp_ag = d_ag * viatico_unit
        texto_recibo = (
            f"<b>RECIBÍ DE TRANSNEA S.A.</b> la suma de <b>$ {imp_ag:,.2f}</b> en concepto de "
            f"viáticos semanales ({d_ag:.1f} días — Sem {com['semana']}) — Comisión N° {com['nro_comision']}/{com['anio']} ({com['lat']}).<br/><br/>"
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

    pie_data = [
        ["TOTAL SEMANAL COMISIÓN ($):", f"$ {float(com['total_comision']):,.2f}"],
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

    story.append(Paragraph(f"TRANSNEA S.A. — RENDICIÓN SEMANAL DE GASTOS N° {ren['nro_rendicion']} (Comisión N° {ren['nro_comision']}/{ren['anio']})", st_title))
    story.append(Spacer(1, 6))

    companeros_str = ", ".join(ren["companeros_json"]) if ren["companeros_json"] else "Sin acompañantes"
    cab_data = [
        ["Del Sr.:", ren["jefe_trabajo"], "LOCALIDAD:", ren["localidad_base"], "FECHA:", f_ren],
        ["Comisión realizada con:", Paragraph(companeros_str, st_small), "", "", "Vehículo:", ren["vehiculo"]],
        ["Resumen Semanal:", Paragraph(str(ren.get("resumen_semanal") or "Cumplimiento semanal programado"), st_small), "", "", "Estado:", "RENDIDA"],
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
        ("SPAN", (1, 2), (3, 2)),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_cab)
    story.append(Spacer(1, 12))

    # Conceptos semanales consolidados (viáticos semanales + peajes/gastos semanales)
    story.append(Paragraph("CONCEPTOS RENDIDOS DE LA SEMANA", st_sub))
    story.append(Spacer(1, 4))

    gastos_rows = [["PERÍODO / REF.", "Comp. Nro.", "CONCEPTO SEMANAL", "IMPORTE ($)"]]
    for g in ren["gastos_detalle_json"]:
        gastos_rows.append([
            str(g.get("periodo", f"Sem {ren.get('semana', '')}")),
            str(g.get("comp_nro", "")),
            str(g.get("concepto", "")),
            f"$ {float(g.get('importe', 0)):,.2f}"
        ])
    gastos_rows.append(["", "", "TOTAL RENDIDO EN LA SEMANA ($)", f"$ {float(ren['total_rendido']):,.2f}"])

    t_gs = Table(gastos_rows, colWidths=[90, 90, 260, 100])
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
    story.append(Paragraph(f"<b>Balance Semanal vs. Anticipo:</b> $ {dif:,.2f} ({estado_dif})", st_small))
    story.append(Paragraph(f"<b>Observaciones / Tareas Cumplidas en la Semana:</b> {ren.get('observaciones') or '-'}", st_small))
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
# 5. INTERFAZ PRINCIPAL DE STREAMLIT
# =============================================================================
cfg = get_config()

st.markdown("""
<div class="header-box">
    <h1>⚡ TRANSNEA S.A. | Flujo Automatizado Semanal: MAPRO ➔ Comisiones ➔ Rendiciones</h1>
    <p>Gestión Semanal Consolidada LAT Zona Centro • PDFs de Comisión y Rendición por Semana</p>
</div>
""", unsafe_allow_html=True)

if not sb:
    st.error("⚠️ Conectá tus credenciales `SUPABASE_URL` y `SUPABASE_KEY` en Secrets para operar con la base de datos en tiempo real.")
    st.stop()

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
    "2️⃣ Disparar Comisiones Semanales",
    "3️⃣ Rendición Semanal de Gastos",
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
            st.caption("Subí tu planilla `LAT - Zona Centro - SEM XX.xlsx` para cargar las tareas de la semana.")
            file_mapro = st.file_uploader("Seleccionar archivo MAPRO (.xlsx)", type=["xlsx"], key="upl_mapro")
            if file_mapro is not None:
                bytes_excel_mapro = file_mapro.getvalue()
                s_imp, a_imp, regs_imp = parsear_excel_mapro(bytes_excel_mapro, sem_sel, anio_sel)
                st.success(f"Detectadas **{len(regs_imp)} filas** de programación para Semana **{s_imp}-{a_imp}**.")
                if st.button("🚀 Guardar MAPRO en Base de Datos", type="primary", use_container_width=True, key="btn_save_mapro_db"):
                    if regs_imp:
                        try:
                            sb.table("programacion_semanal").insert(regs_imp).execute()
                            st.toast("¡MAPRO importada con éxito!", icon="✅")
                            st.rerun()
                        except Exception as e_ins:
                            st.error(f"Error al insertar en Supabase: {e_ins}")
                            st.json(regs_imp[0])

    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Agregar Tarea Semanal Manualmente")
            ca1, ca2, ca3, ca4 = st.columns(4)
            with ca1:
                m_fecha = st.date_input("Fecha de Referencia", value=f_lunes + timedelta(days=1), format="DD/MM/YYYY")
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
                    try:
                        sb.table("programacion_semanal").insert({
                            "id_semana": id_semana_actual,
                            "semana": int(sem_sel),
                            "anio": int(anio_sel),
                            "fecha": m_fecha.isoformat(),
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
                    except Exception as e_add:
                        st.error(f"Error al agregar tarea: {e_add}")

    res_prog = sb.table("programacion_semanal").select("*").eq("id_semana", id_semana_actual).order("fecha").execute()
    df_prog = pd.DataFrame(res_prog.data) if res_prog.data else pd.DataFrame()

    if not df_prog.empty:
        st.markdown("#### 📋 Tareas Programadas en la Semana")
        st.dataframe(
            df_prog[["fecha", "dia", "lat", "tct", "jefe_consigna", "jefe_trabajo", "tarea", "ot_num", "requiere_comision", "comision_generada"]],
            use_container_width=True,
            hide_index=True
        )

        with st.expander("📋 Texto Listo para Correo MAPRO (Copiar y Pegar en Outlook)", expanded=False):
            c_em1, c_em2 = st.columns(2)
            with c_em1:
                st.text_input("Para:", value=DESTINATARIOS_MAPRO_DEFAULT, key="to_mapro")
            with c_em2:
                st.text_input("CC:", value="Oscar Ignacio Orellana", key="cc_mapro")

            st.text_input("Asunto:", value=f"Mapro semana {sem_sel} - LAT - Zona Centro", key="asunto_mapro")

            resumen_por_lat = df_prog.groupby(["lat", "jefe_trabajo", "jefe_consigna"], as_index=False)["tarea"].agg(lambda x: " / ".join(sorted(set(x))))
            lineas_resumen = [
                f"- {r.lat}: {r.tarea} (JT: {r.jefe_trabajo} | JC: {r.jefe_consigna})"
                for r in resumen_por_lat.itertuples()
            ]

            cuerpo_default_mapro = (
                f"Estimados, buenos días.\n\n"
                f"Envío adjunta la MAPRO correspondiente a la Semana {sem_sel} "
                f"(del {f_lunes.strftime('%d/%m/%Y')} al {f_domingo.strftime('%d/%m/%Y')}) de Mantenimiento de LAT Zona Centro.\n\n"
                f"Resumen semanal de trabajos:\n" + "\n".join(lineas_resumen) + "\n\n"
                f"Quedamos a disposición.\nMuchas gracias.\nSaludos cordiales.\n\n"
                f"{cfg['solicitante_default']}\nGerencia de Mantenimiento - Mantenimiento de Líneas"
            )
            st.text_area("Cuerpo del correo:", value=cuerpo_default_mapro, height=200, key="body_mapro")

# =============================================================================
# PESTAÑA 2: DISPARAR COMISIONES SEMANALES Y GENERAR PDFs
# =============================================================================
with tab_comisiones:
    st.markdown(f"### 🧾 Comisiones Semanales Consolidadas — Semana {sem_sel} ({anio_sel})")
    st.caption("Emisión consolidada por semana para cada cuadrilla (total de días en la semana e importe semanal por agente).")

    res_pend = sb.table("programacion_semanal").select("*").eq("id_semana", id_semana_actual).eq("requiere_comision", True).execute()
    df_pend = pd.DataFrame(res_pend.data) if res_pend.data else pd.DataFrame()

    if df_pend.empty:
        st.info("No hay tareas cargadas con comisión pendiente para esta semana. Importá o cargá la MAPRO en la pestaña 1.")
    else:
        grupos = df_pend.groupby(["jefe_trabajo", "lat"], as_index=False).agg({
            "tarea": lambda x: " / ".join(sorted(set(x))),
            "jefe_consigna": "first",
            "ot_num": "first"
        })

        st.markdown("#### ⚡ Cuadrillas Detectadas en la Semana")
        sel_idx = st.selectbox(
            "Seleccioná la cuadrilla / frente semanal:",
            range(len(grupos)),
            format_func=lambda i: f"JT: {grupos.loc[i, 'jefe_trabajo']}  |  LAT: {grupos.loc[i, 'lat']}  |  Tarea: {grupos.loc[i, 'tarea']}"
        )

        g_sel = grupos.loc[sel_idx]

        with st.form("form_generar_comision_semanal"):
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
                cant_dias_sem = st.number_input("Días Totales en la Semana", min_value=0.5, max_value=15.0, value=3.5, step=0.5)
            with fc6:
                val_viatico_com = st.number_input("Valor 1 Viático ($)", min_value=1000.0, value=float(cfg["valor_viatico_diario"]), step=500.0)
            with fc7:
                gastos_ext_com = st.number_input("Gastos Semanales Extra ($)", min_value=0.0, value=0.0, step=500.0)
            with fc8:
                f_emision_com = st.date_input("Fecha de Emisión", value=date.today(), format="DD/MM/YYYY")

            tarea_edit = st.text_input("Tarea Semanal a Realizar", value=str(g_sel["tarea"]).upper())
            obs_com = st.text_input("Observaciones Semanales", placeholder="Opcional")

            st.markdown("##### 👷 Nómina Semanal de Agentes")
            agentes_default = [g_sel["jefe_trabajo"]] if g_sel["jefe_trabajo"] in PERSONAL_MANTENIMIENTO else []
            agentes_sel = st.multiselect(
                "Agentes de la cuadrilla en la semana:",
                options=PERSONAL_MANTENIMIENTO,
                default=agentes_default
            )

            submitted_com = st.form_submit_button("📄 Generar Comisión Semanal en PDF", type="primary", use_container_width=True)

            if submitted_com:
                if not agentes_sel:
                    st.error("Seleccioná al menos 1 agente en la nómina.")
                else:
                    agentes_payload = [
                        {"nombre": a, "funcion": "Oficial", "dias_semana": float(cant_dias_sem)}
                        for a in agentes_sel
                    ]
                    total_calc = (len(agentes_sel) * float(cant_dias_sem) * float(val_viatico_com)) + float(gastos_ext_com)

                    nueva_com = {
                        "nro_comision": int(nro_com_nuevo),
                        "anio": int(anio_sel),
                        "id_semana": id_semana_actual,
                        "semana": int(sem_sel),
                        "fecha_emision": f_emision_com.isoformat(),
                        "inicio_semana": f_lunes.isoformat(),
                        "lat": g_sel["lat"],
                        "ot_num": str(g_sel["ot_num"] or ""),
                        "cuadrilla_base": cuad_base,
                        "vehiculo_1": veh_1,
                        "vehiculo_2": veh_2,
                        "jefe_trabajo": g_sel["jefe_trabajo"],
                        "jefe_consigna": g_sel["jefe_consigna"],
                        "tarea": tarea_edit,
                        "cant_dias": float(cant_dias_sem),
                        "valor_viatico": float(val_viatico_com),
                        "gastos_extra": float(gastos_ext_com),
                        "total_comision": round(total_calc, 2),
                        "agentes_json": agentes_payload,
                        "observaciones": obs_com,
                        "estado": "Emitida",
                        "rendicion_realizada": False
                    }

                    try:
                        sb.table("comisiones_servicio").insert(nueva_com).execute()
                        update_config({"ultimo_nro_comision": int(nro_com_nuevo)})
                        sb.table("programacion_semanal").update({"comision_generada": True}).eq("id_semana", id_semana_actual).eq("jefe_trabajo", g_sel["jefe_trabajo"]).eq("lat", g_sel["lat"]).execute()
                        st.toast(f"¡Comisión Semanal N° {nro_com_nuevo} generada!", icon="🎉")
                        st.rerun()
                    except Exception as e_com:
                        st.error(f"Error guardando comisión: {e_com}")

    st.divider()

    res_coms = sb.table("comisiones_servicio").select("*").eq("id_semana", id_semana_actual).order("nro_comision").execute()
    lista_coms = res_coms.data if res_coms.data else []

    if lista_coms:
        st.markdown(f"#### 📂 Comisiones Semanales Emitidas — Semana {sem_sel} ({len(lista_coms)} comisiones)")
        bloques_texto_mail = []

        for c in lista_coms:
            pdf_bytes = generar_pdf_comision(c, cfg["solicitante_default"], cfg["autorizante_default"])
            apellido_jt = str(c["jefe_trabajo"]).split()[0]
            nombre_pdf = f"SEM {c['semana']} - {c['lat']} - {apellido_jt}.pdf"

            bloques_texto_mail.append(
                f"Comisión {c['nro_comision']} (Semanal)\n"
                f"- Jefe de Trabajo: {c['jefe_trabajo']}\n"
                f"- Tarea: {c['tarea']}\n"
                f"- LAT: {c['lat']}\n"
                f"- Agentes: {len(c['agentes_json'])}\n"
                f"- Días en semana: {str(c['cant_dias']).replace('.', ',')}\n"
            )

            col_c1, col_c2, col_c3 = st.columns([2.5, 1.2, 1.3])
            with col_c1:
                st.write(f"**Comisión N° {c['nro_comision']}** | LAT: `{c['lat']}` | JT: **{c['jefe_trabajo']}** | Agentes: **{len(c['agentes_json'])}** | Total Semanal: **$ {float(c['total_comision']):,.2f}**")
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

        with st.expander("📋 Texto Listo para Correo Contable (Copiar y Pegar en Outlook)", expanded=True):
            ce1, ce2 = st.columns(2)
            with ce1:
                st.text_input("Para (Contable):", value=DESTINATARIOS_CONTABLE_DEFAULT, key="to_contable")
            with ce2:
                st.text_input("CC:", value="Oscar Ignacio Orellana", key="cc_contable")

            st.text_input("Asunto:", value=f"Pedidos de viáticos - Semana {sem_sel}", key="asunto_contable")
            cuerpo_cont_default = (
                f"Estimados, buenos días.\n\n"
                f"Envío los pedidos de comisiones semanales para la semana {sem_sel} de acuerdo con la programación:\n\n"
                + "\n".join(bloques_texto_mail) +
                f"\nQuedo a su disposición, saludos cordiales.\n\n"
                f"{cfg['solicitante_default']}\nGerencia de Mantenimiento - Mantenimiento de Líneas"
            )
            st.text_area("Cuerpo del correo:", value=cuerpo_cont_default, height=240, key="body_contable")

# =============================================================================
# PESTAÑA 3: RENDICIÓN SEMANAL DE GASTOS (POST-TAREA)
# =============================================================================
with tab_rendiciones:
    st.markdown("### 🧾 Rendición Semanal de Gastos (Post-Tarea)")
    st.caption("Liquidación consolidada semanal de viáticos y gastos/peajes de la comisión realizada.")

    res_all_coms = sb.table("comisiones_servicio").select("*").order("nro_comision", desc=True).limit(40).execute()
    coms_disp = res_all_coms.data if res_all_coms.data else []

    if not coms_disp:
        st.info("Aún no hay comisiones emitidas para rendir.")
    else:
        com_sel_ren = st.selectbox(
            "Seleccionar Comisión Semanal a Rendir:",
            coms_disp,
            format_func=lambda c: f"Comisión N° {c['nro_comision']}/{c['anio']} (Sem {c['semana']}) | {c['lat']} | JT: {c['jefe_trabajo']} | Anticipo Semanal: $ {float(c['total_comision']):,.2f} {'✅ Rendida' if c['rendicion_realizada'] else '⏳ Pendiente'}"
        )

        with st.container(border=True):
            st.markdown(f"#### 📝 Rendición Semanal — Comisión N° {com_sel_ren['nro_comision']}/{com_sel_ren['anio']} (Semana {com_sel_ren['semana']})")
            rc1, rc2, rc3, rc4 = st.columns(4)
            with rc1:
                nro_ren_nuevo = st.number_input("N° Rendición", min_value=1, value=int(cfg["ultimo_nro_rendicion"]) + 1, step=1)
            with rc2:
                loc_base_ren = st.text_input("Localidad Base", value="Bqueras.")
            with rc3:
                f_ren_fecha = st.date_input("Fecha Rendición", value=date.today(), format="DD/MM/YYYY")
            with rc4:
                veh_ren = st.text_input("Vehículo Semanal", value=com_sel_ren["vehiculo_1"])

            resumen_sem_txt = st.text_input(
                "Síntesis Semanal de Cumplimiento",
                value=f"Semana {com_sel_ren['semana']}/{com_sel_ren['anio']} — {com_sel_ren['tarea']} en LAT {com_sel_ren['lat']}"
            )

            st.markdown("##### 💵 Conceptos Semanales a Rendir (Viáticos Semanales + Gastos/Peajes de la Semana)")
            cant_ag_ren = len(com_sel_ren["agentes_json"])
            df_gastos_init = pd.DataFrame([
                {
                    "periodo": f"Sem {com_sel_ren['semana']}",
                    "comp_nro": "S/C",
                    "concepto": f"Viáticos semanales (${float(com_sel_ren['valor_viatico']):,.0f}/día x {com_sel_ren['cant_dias']} días x {cant_ag_ren} agentes)",
                    "importe": float(com_sel_ren["total_comision"])
                },
                {
                    "periodo": f"Sem {com_sel_ren['semana']}",
                    "comp_nro": "",
                    "concepto": "Peajes y gastos operativos acumulados en la semana",
                    "importe": 0.0
                }
            ])
            df_gastos_ed = st.data_editor(df_gastos_init, num_rows="dynamic", use_container_width=True, key="ed_gastos_ren_sem")

            total_rendido_calc = float(pd.to_numeric(df_gastos_ed["importe"], errors="coerce").fillna(0).sum())
            dif_calc = total_rendido_calc - float(com_sel_ren["total_comision"])

            km1, km2, km3 = st.columns(3)
            km1.metric("Anticipo Semanal Entregado", f"$ {float(com_sel_ren['total_comision']):,.2f}")
            km2.metric("Total Rendido Semanal", f"$ {total_rendido_calc:,.2f}")
            km3.metric("Diferencia Semanal", f"$ {dif_calc:,.2f}")

            obs_ren = st.text_input("Observaciones de la Rendición Semanal", value=f"Cumplimiento semanal — LAT {com_sel_ren['lat']}")

            if st.button("💾 Guardar Rendición Semanal y Generar PDF", type="primary", use_container_width=True):
                companeros = [a["nombre"] for a in com_sel_ren["agentes_json"] if a["nombre"] != com_sel_ren["jefe_trabajo"]]
                nueva_ren = {
                    "nro_rendicion": int(nro_ren_nuevo),
                    "comision_id": com_sel_ren["id"],
                    "nro_comision": com_sel_ren["nro_comision"],
                    "anio": com_sel_ren["anio"],
                    "fecha_rendicion": f_ren_fecha.isoformat(),
                    "localidad_base": loc_base_ren,
                    "jefe_trabajo": com_sel_ren["jefe_trabajo"],
                    "vehiculo": veh_ren,
                    "companeros_json": companeros,
                    "tramos_json": [],
                    "gastos_detalle_json": df_gastos_ed.to_dict(orient="records"),
                    "total_rendido": round(total_rendido_calc, 2),
                    "diferencia_vs_comision": round(dif_calc, 2),
                    "observaciones": f"{resumen_sem_txt} | {obs_ren}"
                }
                try:
                    sb.table("rendiciones_gastos").insert(nueva_ren).execute()
                    sb.table("comisiones_servicio").update({"rendicion_realizada": True, "estado": "Rendida"}).eq("id", com_sel_ren["id"]).execute()
                    update_config({"ultimo_nro_rendicion": int(nro_ren_nuevo)})
                    st.toast(f"¡Rendición Semanal N° {nro_ren_nuevo} guardada!", icon="✅")
                    st.rerun()
                except Exception as e_ren:
                    st.error(f"Error guardando rendición: {e_ren}")

    res_rens = sb.table("rendiciones_gastos").select("*").order("nro_rendicion", desc=True).limit(20).execute()
    if res_rens.data:
        st.divider()
        st.markdown("#### 📑 Historial de Rendiciones Semanales")
        for r in res_rens.data:
            pdf_ren_bytes = generar_pdf_rendicion(r, cfg["autorizante_default"])
            apellido_r = str(r["jefe_trabajo"]).split()[0]
            fname_r = f"RENDICION_SEM_{r['nro_rendicion']}_COM_{r['nro_comision']}_{apellido_r}.pdf"
            cr1, cr2 = st.columns([3.5, 1.5])
            with cr1:
                st.write(f"**Rendición N° {r['nro_rendicion']}** (Comisión {r['nro_comision']}/{r['anio']}) | JT: **{r['jefe_trabajo']}** | Total Semanal: **$ {float(r['total_rendido']):,.2f}** | Dif: **$ {float(r['diferencia_vs_comision']):,.2f}**")
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

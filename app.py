from __future__ import annotations

import io
import base64
import json
import os
import re
import shutil
import tempfile
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv
from google import genai
from google.genai import types

BASE = Path(__file__).resolve().parent
FUENTES = BASE / "fuentes_permanentes"
INSTRUCCIONES = BASE / "instrucciones" / "instrucciones_juridicas.txt"
TEMP = BASE / "expedientes_temporales"
SALIDAS = BASE / "salidas"
HISTORIAL = SALIDAS / "evaluaciones.xlsx"
FUENTES_REQUERIDAS = [
    "PLANTILLAS cumplimiento.docx",
    "PLANTILLAS DENUNCIAS ACTUALIZADAS.docx",
    "notificaciones mayo.xlsx",
    "CONTADOR DE PLAZOS - TRASU 2026.xlsx",
    "CRITERIOS DE EVALUACION DE CUMPLIMIENTO.xlsx",
    "PAUTAS PAS.xlsx",
]
COLUMNAS = ["Expediente", "Empresa operadora", "Usuario o abonado", "Servicio",
    "Tipo de acto", "Número de resolución o carta", "Fecha de notificación o emisión",
    "Fecha máxima de vencimiento", "Obligación principal", "Resultado",
    "Tipo de incumplimiento", "Subsanación voluntaria", "PAS / NO PAS", "Sustento breve",
    "Párrafo final", "Datos pendientes", "Documentos usados", "Fecha de evaluación"]

load_dotenv(BASE / ".env")
TEMP.mkdir(exist_ok=True); SALIDAS.mkdir(exist_ok=True)

st.set_page_config(page_title="CumpleTRASU", page_icon="⚖️", layout="wide")
st.markdown("""<style>
:root{--navy:#07111f;--card:#101d2d;--cyan:#2dd4bf;--muted:#91a4bb}
.stApp{background:radial-gradient(circle at 50% -20%,#17304b 0,#07111f 45%);color:#e8f0f7}
[data-testid="stHeader"]{background:transparent}.block-container{padding-top:1.25rem;max-width:1600px}
.hero{border:1px solid #24435d;border-radius:16px;padding:18px 24px;background:rgba(12,28,45,.85);margin-bottom:18px}
.brand{font-size:2rem;font-weight:800;color:#fff}.brand b{color:var(--cyan)}.sub{color:#a7bacd}
.tag{float:right;border:1px solid #2d526c;border-radius:999px;padding:7px 12px;color:#9edbd4}
.panel{background:rgba(13,29,46,.9);border:1px solid #203c54;border-radius:14px;padding:14px;margin-bottom:12px}
.light{color:#91a4bb;font-size:.82rem;text-transform:uppercase;letter-spacing:.08em}
.traffic{font-size:1.25rem;font-weight:750;padding:12px;border-radius:10px;text-align:center}
.ok{background:#0b4b3e;color:#9ff6dc}.bad{background:#571e2a;color:#ffc0c7}.wait{background:#574617;color:#ffe49a}
div[data-testid="stFileUploader"]{border:1px dashed #2e6078;border-radius:12px;padding:6px}
.stButton>button{border-radius:9px}.stDownloadButton>button{width:100%;border-radius:9px}
</style>""", unsafe_allow_html=True)
st.markdown('<div class="hero"><span class="tag">OSIPTEL / TRASU</span><div class="brand">Cumple<b>TRASU</b></div><div class="sub">Evaluador Automatizado de Cumplimiento de Resoluciones</div></div>', unsafe_allow_html=True)

def fuente_status() -> tuple[bool, list[str]]:
    faltan = [n for n in FUENTES_REQUERIDAS if not (FUENTES / n).is_file()]
    instrucciones_ok = INSTRUCCIONES.is_file() and INSTRUCCIONES.stat().st_size > 150
    if instrucciones_ok:
        instrucciones_ok = "PEGUE AQUÍ" not in INSTRUCCIONES.read_text("utf-8", errors="ignore")
    if not instrucciones_ok: faltan.append("instrucciones_juridicas.txt (contenido real)")
    return not faltan, faltan

def safe_name(name: str) -> str:
    return re.sub(r"[^\w.() -]", "_", Path(name).name, flags=re.UNICODE)[:180]

def extract_archives(folder: Path) -> list[Path]:
    out=[]
    for f in list(folder.iterdir()):
        try:
            target=folder/(f.stem+"_extraido"); target.mkdir(exist_ok=True)
            if f.suffix.lower()==".zip":
                with zipfile.ZipFile(f) as z:
                    for m in z.infolist():
                        dest=(target/m.filename).resolve()
                        if target.resolve() not in dest.parents and dest != target.resolve(): continue
                        z.extract(m,target)
            elif f.suffix.lower()==".7z":
                import py7zr
                with py7zr.SevenZipFile(f) as z: z.extractall(target)
            elif f.suffix.lower()==".rar":
                import rarfile
                with rarfile.RarFile(f) as z: z.extractall(target)
            else: continue
            out += [p for p in target.rglob("*") if p.is_file()]
        except Exception as e: st.warning(f"No se pudo descomprimir {f.name}: {e}")
    return out

def read_file(path: Path) -> str:
    ext=path.suffix.lower()
    try:
        if ext==".pdf":
            from pypdf import PdfReader
            text="\n".join((p.extract_text() or "") for p in PdfReader(str(path)).pages)
            if len(text.strip())<80:
                try:
                    import pytesseract
                    from pdf2image import convert_from_path
                    text="\n".join(pytesseract.image_to_string(x,lang="spa") for x in convert_from_path(str(path)))
                except Exception: pass
            return text
        if ext==".docx":
            from docx import Document
            d=Document(path); return "\n".join([p.text for p in d.paragraphs]+[" | ".join(c.text for c in r.cells) for t in d.tables for r in t.rows])
        if ext in {".xlsx",".xls"}:
            book=pd.read_excel(path,sheet_name=None,dtype=str)
            return "\n".join(f"[{k}]\n{v.fillna('').to_csv(index=False)}" for k,v in book.items())
        if ext==".csv": return pd.read_csv(path,dtype=str).fillna("").to_csv(index=False)
        if ext in {".txt",".md"}: return path.read_text("utf-8",errors="ignore")
        if ext in {".png",".jpg",".jpeg",".tif",".tiff",".bmp"}:
            try:
                import pytesseract
                from PIL import Image
                text=pytesseract.image_to_string(Image.open(path),lang="spa")
                if text.strip(): return text
            except Exception: pass
            return vision_ocr(path)
    except Exception as e: return f"[Error al leer {path.name}: {e}]"
    return ""

def vision_ocr(path: Path) -> str:
    """OCR cloud fallback for scanned images, including TIFF converted to JPEG."""
    try:
        from PIL import Image, ImageSequence
        source=Image.open(path)
        parts=[types.Part.from_text(text="Transcribe fielmente todo el texto jurídico visible, página por página. No resumas ni inventes.")]
        for frame in list(ImageSequence.Iterator(source))[:20]:
            image=frame.convert("RGB"); image.thumbnail((1800,1800))
            buf=io.BytesIO(); image.save(buf,format="JPEG",quality=85)
            parts.append(types.Part.from_bytes(data=buf.getvalue(),mime_type="image/jpeg"))
        response=get_client(get_api_key()).models.generate_content(
            model=os.getenv("GEMINI_MODEL","gemini-2.0-flash"),contents=parts)
        return response.text
    except Exception as e:
        return f"[OCR no disponible para {path.name}: {e}]"

def classify(name: str, text: str) -> str:
    s=(name+" "+text[:3000]).lower()
    for key,label in [("sara","SARA"),("sar","SAR"),("sap","SAP"),("denuncia","Denuncia"),("carta","Carta"),("primera instancia","Resolución de primera instancia"),("trasu","Resolución TRASU")]:
        if key in s: return label
    return "No identificado"

def exact_notification(expediente: str) -> str | None:
    path=FUENTES/"notificaciones mayo.xlsx"
    for _,df in pd.read_excel(path,sheet_name=None,dtype=str).items():
        cols={str(c).strip().upper():c for c in df.columns}
        date_key="FEC_NOT_EMP_ELE_TEXTO" if "FEC_NOT_EMP_ELE_TEXTO" in cols else ("FEC_NOT_EMP_ELE" if "FEC_NOT_EMP_ELE" in cols else None)
        if "NRO_EXPEDIENTE" in cols and date_key:
            values=df[cols["NRO_EXPEDIENTE"]].fillna("").astype(str).str.strip()
            hit=df.loc[values==expediente.strip(),cols[date_key]]
            if not hit.empty and pd.notna(hit.iloc[0]): return str(hit.iloc[0]).strip()
    return None

def identify_exact_expediente(context: str) -> str | None:
    """Resolve filename-safe variants to one exact value stored in NRO_EXPEDIENTE."""
    normalized_context=re.sub(r"[^A-Z0-9]","",context.upper())
    candidates=[]
    path=FUENTES/"notificaciones mayo.xlsx"
    for _,df in pd.read_excel(path,sheet_name=None,dtype=str).items():
        cols={str(c).strip().upper():c for c in df.columns}
        if "NRO_EXPEDIENTE" not in cols: continue
        for value in df[cols["NRO_EXPEDIENTE"]].dropna().astype(str).str.strip().unique():
            normalized=re.sub(r"[^A-Z0-9]","",value.upper())
            if len(normalized)>=10 and normalized in normalized_context:
                candidates.append(value)
    unique=list(dict.fromkeys(candidates))
    return unique[0] if len(unique)==1 else None

def parse_trasu_name(names: str) -> str | None:
    compact=re.sub(r"[^A-Z0-9]","",names.upper())
    match=re.search(r"(\d{7})(20\d{2})TRASUSTRA",compact)
    return f"{match.group(1)}-{match.group(2)}/TRASU/ST-RA" if match else None

def source_text(name: str, limit=50000) -> str:
    return read_file(FUENTES/name)[:limit]

def calculate_due(notification: str, context: str) -> tuple[str | None, str | None]:
    """Extract explicit calculator inputs, then reproduce Calculadora libre exactly.

    Returns (fecha_vencimiento, motivo_error). motivo_error is set whenever the
    calculation could not be completed, so the UI can explain the real cause
    instead of a generic message.
    """
    key=get_api_key()
    if not key: return None,"no hay una clave de Gemini configurada"
    try:
        r=get_client(key).models.generate_content(
            model=os.getenv("GEMINI_MODEL","gemini-2.0-flash"),
            contents=context[:60000],
            config=types.GenerateContentConfig(
                system_instruction="Extrae SOLO datos expresos de la resolución para usar la pestaña Calculadora libre. Devuelve JSON: {numero_dias: integer|null, tipo_dias: 'Habiles'|'Calendario'|null, ubicacion: 'Lima'|'Otra'|null, evidencia: string}. No infieras un plazo ausente.",
                response_mime_type="application/json"))
    except Exception as e:
        return None,f"la IA no pudo procesar la solicitud ({e})"
    try:
        params=json.loads(r.text)
    except Exception as e:
        return None,f"la IA devolvió una respuesta no válida ({e})"
    days=params.get("numero_dias"); kind=params.get("tipo_dias"); location=params.get("ubicacion")
    if not isinstance(days,int) or days<0 or kind not in {"Habiles","Calendario"}:
        return None,"no se identificó en la resolución un plazo expreso en días (hábiles o calendario)"
    start=pd.to_datetime(notification,errors="coerce")
    if pd.isna(start): return None,f"la fecha de notificación '{notification}' no se pudo interpretar"
    if kind=="Calendario": return (start+timedelta(days=days)).strftime("%Y-%m-%d"),None
    try:
        holidays_book=pd.read_excel(FUENTES/"CONTADOR DE PLAZOS - TRASU 2026.xlsx",sheet_name="No laborables (2)",header=None)
    except Exception as e:
        return None,f"no se pudo leer la hoja 'No laborables (2)' de CONTADOR DE PLAZOS - TRASU 2026.xlsx ({e})"
    col=1 if location=="Lima" else 2
    if col>=holidays_book.shape[1]: return None,"CONTADOR DE PLAZOS - TRASU 2026.xlsx no tiene la columna de días no laborables esperada"
    holidays=pd.to_datetime(holidays_book.iloc[:,col],errors="coerce").dropna().dt.normalize().unique()
    offset=pd.offsets.CustomBusinessDay(n=days,weekmask="Mon Tue Wed Thu Fri",holidays=list(holidays))
    return (start.normalize()+offset).strftime("%Y-%m-%d"),None

def get_api_key() -> str | None:
    try: return st.secrets.get("GOOGLE_API_KEY") or os.getenv("GOOGLE_API_KEY")
    except Exception: return os.getenv("GOOGLE_API_KEY")

@st.cache_resource(show_spinner=False)
def get_client(api_key: str) -> genai.Client:
    """Reuse a single Gemini client; creating one per call closes shared connections mid-request."""
    return genai.Client(api_key=api_key)

def ai_evaluate(payload: dict[str,Any]) -> dict[str,Any]:
    template="PLANTILLAS cumplimiento.docx" if payload["tipo_acto"]=="Resolución TRASU" else "PLANTILLAS DENUNCIAS ACTUALIZADAS.docx"
    sources={"instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore")[:50000],"plantilla":source_text(template),
        "criterios":source_text("CRITERIOS DE EVALUACION DE CUMPLIMIENTO.xlsx"),"pautas_pas":source_text("PAUTAS PAS.xlsx")}
    schema={"ficha":{"expediente":"","empresa_operadora":"","usuario_abonado":"","servicio":"","tipo_acto":"","numero_acto":"","fecha_notificacion_emision":"","fecha_vencimiento":"","obligacion_principal":"","medios_probatorios":[]},"trazabilidad":[],"evaluacion_juridica":{"checklist":[],"sustento_breve":"","tipo_incumplimiento":""},"resultado":"Cumplió|Incumplió|Inejecutable","subsanacion_voluntaria":"aplica|no aplica|no corresponde","clasificacion":"PAS|NO PAS","parrafo_final":"","datos_pendientes":[]}
    system="Eres analista jurídico de OSIPTEL. Usa SOLO la evidencia y fuentes entregadas. No inventes fechas, pruebas ni conclusiones. Lo faltante es 'No identificado' o 'Pendiente de verificación'. Devuelve JSON válido conforme al esquema."
    user=json.dumps({"esquema":schema,"caso":payload,"fuentes":sources},ensure_ascii=False)
    r=get_client(get_api_key()).models.generate_content(
        model=os.getenv("GEMINI_MODEL","gemini-2.0-flash"),contents=user,
        config=types.GenerateContentConfig(system_instruction=system,response_mime_type="application/json"))
    return json.loads(r.text)

def regenerate_paragraph(result: dict[str,Any]) -> str:
    context={"ficha":result.get("ficha",{}),"evaluacion_juridica":result.get("evaluacion_juridica",{}),"resultado":result.get("resultado"),"subsanacion_voluntaria":result.get("subsanacion_voluntaria"),"clasificacion":result.get("clasificacion"),"datos_pendientes":result.get("datos_pendientes",[]),"instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore")[:40000]}
    response=get_client(get_api_key()).models.generate_content(
        model=os.getenv("GEMINI_MODEL","gemini-2.0-flash"),contents=json.dumps(context,ensure_ascii=False),
        config=types.GenerateContentConfig(system_instruction="Regenera únicamente el párrafo jurídico final con los datos aportados. No inventes ni completes faltantes. Devuelve JSON {parrafo_final:string}.",response_mime_type="application/json"))
    return json.loads(response.text)["parrafo_final"]

def to_row(result: dict, documents: list[str]) -> dict:
    f=result.get("ficha",{}); e=result.get("evaluacion_juridica",{})
    return dict(zip(COLUMNAS,[f.get("expediente",""),f.get("empresa_operadora",""),f.get("usuario_abonado",""),f.get("servicio",""),f.get("tipo_acto",""),f.get("numero_acto",""),f.get("fecha_notificacion_emision",""),f.get("fecha_vencimiento",""),f.get("obligacion_principal",""),result.get("resultado",""),e.get("tipo_incumplimiento",""),result.get("subsanacion_voluntaria",""),result.get("clasificacion",""),e.get("sustento_breve",""),result.get("parrafo_final",""),"; ".join(map(str,result.get("datos_pendientes",[]))),"; ".join(documents),datetime.now().strftime("%Y-%m-%d %H:%M")]))

def excel_bytes(row: dict | None=None) -> bytes:
    existing=pd.read_excel(HISTORIAL,dtype=str) if HISTORIAL.exists() else pd.DataFrame(columns=COLUMNAS)
    if row is not None: existing=pd.concat([existing,pd.DataFrame([row])],ignore_index=True)
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w: existing.to_excel(w,index=False,sheet_name="Evaluaciones")
    return b.getvalue()

for k,v in {"result":None,"docs":[],"texts":{},"notice":None,"due":None}.items(): st.session_state.setdefault(k,v)

left,center,right=st.columns([1.05,2.25,1.05],gap="large")
with left:
    st.markdown('<div class="light">Expediente particular</div>',unsafe_allow_html=True)
    uploads=st.file_uploader("Cargar expediente",accept_multiple_files=True,type=["pdf","docx","xlsx","xls","csv","txt","png","jpg","jpeg","tif","tiff","zip","rar","7z"])
    if uploads:
        st.session_state.docs=[u.name for u in uploads]
        for u in uploads:
            kind=classify(u.name,""); st.caption(f"✓ {u.name} · {kind}")
    ok_sources,missing=fuente_status()
    st.divider(); st.markdown('<div class="light">Fuentes permanentes</div>',unsafe_allow_html=True)
    st.success("6 fuentes + instrucciones disponibles") if ok_sources else st.error(f"Faltan {len(missing)} fuentes")
    if missing:
        with st.expander("Ver faltantes"): st.write("\n".join(f"• {x}" for x in missing))
    analyze=st.button("✦ Analizar expediente con IA",type="primary",use_container_width=True,disabled=not uploads)

if analyze:
    if not ok_sources: st.error("La herramienta no puede emitir evaluación final porque no tiene acceso a las fuentes permanentes")
    elif not get_api_key(): st.error("Configure GOOGLE_API_KEY en .env o en los secretos de Streamlit.")
    else:
        with st.spinner("Procesando y contrastando el expediente..."):
            case_dir=Path(tempfile.mkdtemp(prefix="caso_",dir=TEMP))
            try:
                paths=[]
                for u in uploads:
                    p=case_dir/safe_name(u.name); p.write_bytes(u.getbuffer()); paths.append(p)
                paths+=extract_archives(case_dir)
                texts={p.name:read_file(p) for p in paths if p.suffix.lower() not in {".zip",".rar",".7z"}}
                combined="\n\n".join(f"### {n}\n{t}" for n,t in texts.items())
                tipos_detectados=[classify(n,t) for n,t in texts.items()]; tipo=next((x for x in tipos_detectados if x!="No identificado"),"No identificado")
                searchable="\n".join(u.name for u in uploads)
                expediente=parse_trasu_name(searchable) or identify_exact_expediente(searchable) or "No identificado"
                notice=None; due=None
                if tipo=="Resolución TRASU":
                    notice=exact_notification(expediente) if expediente!="No identificado" else None
                    if not notice: st.error("No se puede emitir evaluación final porque no se encontró coincidencia exacta del expediente en notificaciones mayo.xlsx")
                    else:
                        due,due_error=calculate_due(notice,combined)
                        if not due: st.error(f"No se puede emitir evaluación final porque no se pudo calcular el vencimiento con CONTADOR DE PLAZOS - TRASU 2026.xlsx: {due_error}")
                if tipo!="Resolución TRASU" or (notice and due):
                    payload={"tipo_acto":tipo,"expediente_detectado":expediente,"fecha_verificada":notice or "No identificado","fecha_vencimiento":due or "Pendiente de verificación","documentos":texts}
                    st.session_state.result=ai_evaluate(payload); st.session_state.texts=texts
            except Exception as e: st.error(f"No se pudo completar el análisis: {e}")
            finally: shutil.rmtree(case_dir,ignore_errors=True)

r=st.session_state.result
with center:
    st.markdown('<div class="light">Ficha editable del expediente</div>',unsafe_allow_html=True)
    f=(r or {}).get("ficha",{})
    a,b=st.columns(2)
    expediente=a.text_input("Expediente",f.get("expediente","No identificado")); empresa=b.text_input("Empresa operadora",f.get("empresa_operadora","No identificado"))
    usuario=a.text_input("Usuario o abonado",f.get("usuario_abonado","No identificado")); servicio=b.text_input("Servicio",f.get("servicio","No identificado"))
    tipo=a.text_input("Tipo de acto",f.get("tipo_acto","No identificado")); numero=b.text_input("Número de resolución o carta",f.get("numero_acto","No identificado"))
    notif=a.text_input("Fecha de notificación o emisión",f.get("fecha_notificacion_emision","No identificado")); vence=b.text_input("Fecha máxima de vencimiento",f.get("fecha_vencimiento","Pendiente de verificación"))
    obligacion=st.text_area("Obligación principal",f.get("obligacion_principal","No identificado"),height=90)
    st.markdown("**Medios probatorios**")
    medios=f.get("medios_probatorios",[]) if r else []
    st.write("\n".join(f"• {x}" for x in medios) if medios else "No identificado")
    with st.expander("Checklist jurídico",expanded=True):
        checks=(r or {}).get("evaluacion_juridica",{}).get("checklist",[])
        st.write("\n".join(f"• {x}" for x in checks) if checks else "Pendiente de evaluación")
    with st.expander("Trazabilidad"):
        trace=(r or {}).get("trazabilidad",[])
        st.json(trace) if trace else st.write("Pendiente de evaluación")

with right:
    st.markdown('<div class="light">Dictamen</div>',unsafe_allow_html=True)
    resultado=(r or {}).get("resultado","Pendiente")
    css="ok" if resultado=="Cumplió" else "bad" if resultado=="Incumplió" else "wait"
    st.markdown(f'<div class="traffic {css}">{resultado}</div>',unsafe_allow_html=True)
    st.metric("Subsanación voluntaria",(r or {}).get("subsanacion_voluntaria","Pendiente"))
    st.metric("Clasificación",(r or {}).get("clasificacion","Pendiente"))
    st.markdown("**Alertas**")
    pendientes=(r or {}).get("datos_pendientes",[])
    if pendientes:
        for x in pendientes: st.warning(str(x))
    else: st.caption("Sin alertas disponibles")

st.divider(); st.markdown('<div class="light">Párrafo final</div>',unsafe_allow_html=True)
paragraph=st.text_area("Resultado listo para copiar",(r or {}).get("parrafo_final","La evaluación aparecerá aquí cuando se complete el análisis."),height=150,label_visibility="collapsed")
components.html(f"""<button onclick='navigator.clipboard.writeText(document.getElementById("p").textContent)' style='background:#15324a;color:#e8f0f7;border:1px solid #2d526c;border-radius:8px;padding:8px 15px;cursor:pointer'>Copiar párrafo</button><span id='p' style='display:none'>{paragraph.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')}</span>""",height=45)
c1,c2,c3,c4=st.columns(4)
with c1:
    if st.button("Guardar evaluación",use_container_width=True,disabled=not bool(r)):
        f.update({"expediente":expediente,"empresa_operadora":empresa,"usuario_abonado":usuario,"servicio":servicio,"tipo_acto":tipo,"numero_acto":numero,"fecha_notificacion_emision":notif,"fecha_vencimiento":vence,"obligacion_principal":obligacion}); r["parrafo_final"]=paragraph
        row=to_row(r,st.session_state.docs); data=excel_bytes(row); HISTORIAL.write_bytes(data); st.success("Evaluación guardada")
with c2: st.download_button("Exportar a Excel",excel_bytes(to_row(r,st.session_state.docs)) if r else excel_bytes(),"CumpleTRASU_evaluacion.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
with c3:
    if st.button("Regenerar párrafo",use_container_width=True,disabled=not bool(r)):
        f.update({"expediente":expediente,"empresa_operadora":empresa,"usuario_abonado":usuario,"servicio":servicio,"tipo_acto":tipo,"numero_acto":numero,"fecha_notificacion_emision":notif,"fecha_vencimiento":vence,"obligacion_principal":obligacion,"medios_probatorios":medios})
        try:
            r["parrafo_final"]=regenerate_paragraph(r); st.session_state.result=r; st.rerun()
        except Exception as e: st.error(f"No se pudo regenerar el párrafo: {e}")
with c4:
    if st.button("Limpiar expediente",use_container_width=True):
        for k in ["result","docs","texts","notice","due"]: st.session_state[k]=None if k=="result" else ([] if k=="docs" else {})
        st.rerun()

st.divider(); st.markdown("### Casos evaluados")
if HISTORIAL.exists(): st.dataframe(pd.read_excel(HISTORIAL,dtype=str),use_container_width=True,hide_index=True)
else: st.caption("Aún no hay evaluaciones guardadas.")
st.caption("CumpleTRASU asiste el análisis jurídico; la revisión profesional y la integridad de las fuentes siguen siendo obligatorias.")

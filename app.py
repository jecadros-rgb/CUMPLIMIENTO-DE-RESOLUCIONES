from __future__ import annotations

import io
import base64
import hashlib
import json
import os
import re
import shutil
import tempfile
import unicodedata
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
CRITERIOS_INSTRUCCION = BASE / "instrucciones" / "criterios_evaluacion_obligatorios.txt"
TEMP = BASE / "expedientes_temporales"
SALIDAS = BASE / "salidas"
HISTORIAL = SALIDAS / "evaluaciones.xlsx"
FUENTES_REQUERIDAS = [
    "PLANTILLAS cumplimiento.docx",
    "PLANTILLAS DENUNCIAS ACTUALIZADAS.docx",
    "notificaciones mayo.xlsx",
    "CONTADOR DE PLAZOS - TRASU 2026.xlsx",
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

def get_api_key() -> str | None:
    try: return st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    except Exception: return os.getenv("GEMINI_API_KEY")

def gemini_client():
    return genai.Client(api_key=get_api_key())

def gemini_text(system: str, user: Any, json_mode: bool=False) -> str:
    config=types.GenerateContentConfig(system_instruction=system,max_output_tokens=32768)
    if json_mode: config.response_mime_type="application/json"
    client=gemini_client()
    errors=[]
    # Flash-Lite first: it is both far cheaper and has a much higher free-tier
    # daily quota than 3.5-flash, so prefer it in every call, not only OCR.
    models=("gemini-3.1-flash-lite","gemini-3.5-flash")
    for model in models:
        try:
            response=client.models.generate_content(model=model,contents=user,config=config)
            candidates=response.candidates or []
            finish=str(candidates[0].finish_reason) if candidates else "SIN_CANDIDATOS"
            text=response.text if candidates and candidates[0].content and candidates[0].content.parts else None
            if not text: raise RuntimeError(f"El modelo {model} no devolvió contenido utilizable (motivo: {finish}).")
            return text
        except Exception as e:
            errors.append(f"{model}: {e}")
    raise RuntimeError(" | ".join(errors))

def parse_json_response(text: str) -> Any:
    """Gemini's JSON mode can still wrap output in markdown fences or append trailing text; recover the first valid value."""
    cleaned=text.strip()
    if cleaned.startswith("```"):
        cleaned=re.sub(r"^```[a-zA-Z]*\n?","",cleaned)
        cleaned=re.sub(r"```\s*$","",cleaned).strip()
    return json.JSONDecoder().raw_decode(cleaned)[0]

def fuente_status() -> tuple[bool, list[str]]:
    faltan = [n for n in FUENTES_REQUERIDAS if not (FUENTES / n).is_file()]
    instrucciones_ok = INSTRUCCIONES.is_file() and INSTRUCCIONES.stat().st_size > 150
    if instrucciones_ok:
        instrucciones_ok = "PEGUE AQUÍ" not in INSTRUCCIONES.read_text("utf-8", errors="ignore")
    if not instrucciones_ok: faltan.append("instrucciones_juridicas.txt (contenido real)")
    if not CRITERIOS_INSTRUCCION.is_file() or CRITERIOS_INSTRUCCION.stat().st_size < 500:
        faltan.append("criterios_evaluacion_obligatorios.txt")
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
                from PIL import Image, ImageSequence
                source=Image.open(path)
                text="\n".join(pytesseract.image_to_string(frame.convert("RGB"),lang="spa")
                               for frame in ImageSequence.Iterator(source))
                legal_markers=("resuelve","declarar fundad","declara fundad","artículo 1","articulo 1","sentido de la resolución")
                if text.strip() and any(x in text.lower() for x in legal_markers): return text
            except Exception: pass
            return vision_ocr(path)
    except Exception as e: return f"[Error al leer {path.name}: {e}]"
    return ""

def vision_ocr(path: Path) -> str:
    """OCR cloud fallback for scanned images, including multi-page TIFF converted to JPEG."""
    try:
        from PIL import Image, ImageSequence
        source=Image.open(path)
        content=["Transcribe fielmente todo el texto jurídico visible en cada una de las páginas de imagen adjuntas, en orden, marcando cada una con '--- Página N ---'. No resumas, no omitas ninguna página ni inventes contenido."]
        # IMPORTANT: do not wrap this in list(...) first. ImageSequence.Iterator
        # re-seeks the SAME underlying image object on every step; materializing
        # a list before converting each frame leaves every entry pointing at the
        # last frame's data. Convert to RGB immediately, inside the loop, so each
        # frame's pixels are captured before the iterator advances.
        for i,frame in enumerate(ImageSequence.Iterator(source)):
            if i>=20: break
            image=frame.convert("RGB"); image.thumbnail((1800,1800))
            buf=io.BytesIO(); image.save(buf,format="JPEG",quality=85)
            content.append(types.Part.from_bytes(data=buf.getvalue(),mime_type="image/jpeg"))
        return gemini_text("Eres un sistema OCR jurídico preciso.",content)
    except Exception as e:
        return f"[OCR no disponible para {path.name}: {e}]"

def classify(name: str, text: str) -> str:
    s=(name+" "+text[:3000]).lower()
    for key,label in [("sara","SARA"),("sar","SAR"),("sap","SAP"),("denuncia","Denuncia"),("carta","Carta"),("primera instancia","Resolución de primera instancia"),("trasu","Resolución TRASU")]:
        if key in s: return label
    return "No identificado"

def extract_resolutive_part(documents: dict[str,str]) -> str | None:
    """Extract the operative section; obligations may not be inferred elsewhere."""
    candidates=[]
    for name,text in documents.items():
        if classify(name,text)!="Resolución TRASU": continue
        upper=unicodedata.normalize("NFD",text.upper())
        upper="".join(c for c in upper if unicodedata.category(c)!="Mn")
        anchors=[]
        for pattern in (r"\bHA\s+RESUELTO\b",r"\bSE\s+RESUELVE\b",r"\bRESUELVE\s*:",r"\bPARTE\s+RESOLUTIVA\b",
                        r"\bARTICULO\s+(?:PRIMERO|1)\b",r"\bART\.\s*1\b"):
            match=re.search(pattern,upper)
            if match: anchors.append(match.start())
        if anchors:
            start=min(anchors); section=text[start:start+18000]
            if re.search(r"DECLARAR\s+(?:EL\s+RECLAMO\s+)?FUNDAD",section,re.I):
                candidates.append(f"### {name} — PARTE RESOLUTIVA\n{section}")
            else:
                candidates.append(f"### {name} — PARTE RESOLUTIVA\n{section}")
            continue
        # OCR sometimes omits the heading but preserves the operative declaration.
        match=re.search(r"(?:DECLARAR|DECLARA|SE\s+DECLARA)[^.\n]{0,180}?FUNDAD[OA]",upper)
        if match: candidates.append(f"### {name} — DECLARACIÓN FUNDADA\n{text[max(0,match.start()-800):match.start()+12000]}")
        elif re.search(r"SENTIDO\s+DE\s+LA\s+RESOLUCION[^.\n]{0,100}FUNDAD[OA]",upper):
            # Some official copies place the operative articles on the final pages.
            candidates.append(f"### {name} — PÁGINAS FINALES DE RESOLUCIÓN FUNDADA\n{text[-18000:]}")
        elif len(text.strip())>500:
            # Last-resort candidate is restricted to the end of the resolution,
            # never to letters, briefs or evidence from other documents.
            candidates.append(f"### {name} — CANDIDATO DE PARTE RESOLUTIVA (PÁGINAS FINALES)\n{text[-18000:]}")
    return "\n\n".join(candidates) if candidates else None

def exact_notification(expediente: str) -> str | None:
    path=FUENTES/"notificaciones mayo.xlsx"
    for _,df in pd.read_excel(path,sheet_name=None,dtype=str).items():
        cols={str(c).strip().upper():c for c in df.columns}
        date_key="FEC_NOT_EMP_ELE_TEXTO" if "FEC_NOT_EMP_ELE_TEXTO" in cols else ("FEC_NOT_EMP_ELE" if "FEC_NOT_EMP_ELE" in cols else None)
        if "NRO_EXPEDIENTE" in cols and date_key:
            values=df[cols["NRO_EXPEDIENTE"]].fillna("").astype(str).str.strip()
            hit=df.loc[values==expediente.strip(),cols[date_key]]
            if not hit.empty and pd.notna(hit.iloc[0]):
                raw=str(hit.iloc[0]).strip()
                # The source column stores a full timestamp (00:00:00); keep only the date.
                parsed=pd.to_datetime(raw,errors="coerce")
                return parsed.strftime("%d/%m/%Y") if pd.notna(parsed) else raw
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

def parse_excel_date(value: Any) -> pd.Timestamp:
    """Accept displayed dates, timestamps and Excel serial date numbers."""
    raw=str(value).strip()
    if re.fullmatch(r"\d+(?:\.0+)?",raw):
        serial=float(raw)
        if 20000 <= serial <= 80000:
            return pd.Timestamp("1899-12-30")+pd.to_timedelta(serial,unit="D")
    parsed=pd.to_datetime(raw,dayfirst=True,errors="coerce")
    if pd.isna(parsed): raise ValueError(f"fecha de notificación no reconocida: {raw}")
    return pd.Timestamp(parsed)

def _legal_tokens(value: Any) -> set[str]:
    text=unicodedata.normalize("NFD",str(value).lower())
    text="".join(c for c in text if unicodedata.category(c)!="Mn")
    return {x for x in re.findall(r"[a-z0-9]{4,}",text) if x not in {
        "para","como","esta","este","desde","hasta","sobre","entre","donde",
        "empresa","operadora","resolucion","trasu","usuario","cumplimiento"}}

def relevant_excel_rules(name: str, case_context: str, limit: int=28) -> str:
    """Retrieve applicable rows instead of truncating a whole legal workbook."""
    book=pd.read_excel(FUENTES/name,sheet_name=None,dtype=str,header=None)
    query=_legal_tokens(case_context)
    ranked=[]
    for sheet,df in book.items():
        for idx,row in df.fillna("").iterrows():
            cells=[str(x).strip() for x in row.tolist() if str(x).strip()]
            if not cells: continue
            line=" | ".join(cells)
            tokens=_legal_tokens(line)
            score=len(query & tokens)
            low=line.lower()
            # Rules that govern every case must never disappear from context.
            general=any(k in low for k in (
                "cese total","reversión total","subsanación voluntaria",
                "análisis por cada mandato","conclusión es por resolución",
                "cumplimiento parcial","registro y activación de los descuentos",
                "descuento recurrente","programación","no se cuente con ello, hay infracción"))
            if general: score+=100
            if score: ranked.append((score,sheet,int(idx)+1,line))
    ranked.sort(key=lambda x:(-x[0],x[1],x[2]))
    selected=ranked[:limit]
    return "\n".join(f"[{name} / {s} / fila {r}] {line}" for _,s,r,line in selected)

def legal_sources(case_context: str, template: str) -> dict[str,str]:
    return {
        "instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore"),
        "criterios_evaluacion_instruccion_obligatoria":CRITERIOS_INSTRUCCION.read_text("utf-8",errors="ignore"),
        "plantilla_aplicable":source_text(template,80000),
        "pautas_pas_aplicables":relevant_excel_rules("PAUTAS PAS.xlsx",case_context,26),
    }

def calculate_due(notification: str, context: str) -> tuple[str,str] | None:
    """Determine the legal term without AI and reproduce Calculadora libre."""
    try:
        normalized=unicodedata.normalize("NFD",context.lower())
        normalized="".join(c for c in normalized if unicodedata.category(c)!="Mn")
        start=parse_excel_date(notification)
        recurring_discount=("descuento" in normalized and any(k in normalized for k in (
            "descuento recurrente","ajustes recurrentes","seis (6) meses","seis meses",
            "por 6 meses","periodo de 6 meses","periodo total de seis",
            "periodo de seis","meses pendientes","meses restantes")))
        explicit=re.findall(r"(?:plazo|termino)[^.\n]{0,160}?(?:\(|\b)(\d{1,3})\)?\s*dias?\s*habiles",normalized,re.I)
        explicit_days={int(x) for x in explicit}
        if recurring_discount:
            days=10
        elif len(explicit_days)==1:
            days=next(iter(explicit_days))
        elif 10 in explicit_days:
            days=10
        elif re.search(r"(?:plazo[^.]{0,80})?(?:un\s*\(1\)|1)\s*mes",normalized):
            return (start+pd.DateOffset(months=1)).strftime("%d/%m/%Y"),"1 mes"
        else:
            days=10
        term=f"{days} días hábiles"
        holidays_book=pd.read_excel(FUENTES/"CONTADOR DE PLAZOS - TRASU 2026.xlsx",sheet_name="No laborables (2)",header=None)
        if holidays_book.shape[1] < 2: raise ValueError("la hoja No laborables (2) no contiene la columna Lima")
        holidays=pd.to_datetime(holidays_book.iloc[:,1],dayfirst=True,errors="coerce").dropna().dt.normalize().unique()
        offset=pd.offsets.CustomBusinessDay(n=days,weekmask="Mon Tue Wed Thu Fri",holidays=list(holidays))
        return (start.normalize()+offset).strftime("%d/%m/%Y"),term
    except Exception as e:
        raise ValueError(f"error del contador de plazos: {e}") from e

def deterministic_paragraph(result: dict[str,Any], extraction: dict[str,Any]) -> str:
    """Last-resort grounded drafting; never leaves a completed case blank."""
    f=result.get("ficha") if isinstance(result.get("ficha"),dict) else {}
    acto=f.get("tipo_acto") or "acto evaluado"
    notif=f.get("fecha_notificacion_emision") or "no identificada"
    plazo=f.get("plazo_cumplimiento") or "no identificado"
    vence=f.get("fecha_vencimiento") or "no identificada"
    obligacion=f.get("obligacion_principal") or "la obligación ordenada"
    pruebas=[]
    for item in (extraction.get("medios_probatorios") or []):
        if not isinstance(item,dict): continue
        doc=item.get("documento") or "documento aportado"
        fecha=f" de fecha {item.get('fecha')}" if item.get("fecha") else ""
        hecho=item.get("hecho_acreditado") or "sin hecho acreditado identificado"
        estado=item.get("estado") or "no acreditado"
        pruebas.append(f"{doc}{fecha}: {hecho} (estado: {estado})")
    matrices=[]
    for item in (extraction.get("matriz_cumplimiento") or []):
        if isinstance(item,dict):
            matrices.append(f"{item.get('componente') or 'obligación'}: {item.get('estado') or 'no acreditado'}; {item.get('sustento') or 'sin sustento adicional'}")
    contraste="; ".join(pruebas) if pruebas else "no se identificaron medios probatorios objetivos suficientes"
    matriz="; ".join(matrices) if matrices else "no se acreditó documentalmente la ejecución de cada componente"
    resultado=result.get("resultado") or "Pendiente"
    subs=result.get("subsanacion_voluntaria") or "no corresponde"
    clas=result.get("clasificacion") or "Pendiente"
    return (f"El {acto} fue notificado el {notif}, con un plazo de {plazo}, que vencía el {vence}. "
            f"La obligación principal consistía en {obligacion}. Al respecto, la documentación examinada acredita lo siguiente: {contraste}. "
            f"Del contraste individual de las obligaciones se obtiene: {matriz}. En consecuencia, el resultado de la evaluación es {resultado} y la clasificación es {clas}. "
            f"Respecto de la subsanación voluntaria, {subs}, conforme al cese y reversión efectivamente acreditados en el expediente.")

def ai_evaluate(payload: dict[str,Any]) -> dict[str,Any]:
    template="PLANTILLAS cumplimiento.docx" if payload["tipo_acto"]=="Resolución TRASU" else "PLANTILLAS DENUNCIAS ACTUALIZADAS.docx"
    case_context=json.dumps(payload,ensure_ascii=False)
    sources=legal_sources(case_context,template)
    extraction_schema={"acto":{"numero":"","fecha":"","mandato_textual":""},"obligacion_extraida_parte_resolutiva":{"texto":"","articulo_numeral":"","declara_fundado":"si|no|no_identificado"},"obligaciones":[{"componente":"","periodo":"","plazo_expreso":"","prueba_exigible":""}],"medios_probatorios":[{"documento":"","fecha":"","hecho_acreditado":"","cita":"","estado":"ejecutado|programado|en_curso|no_acreditado"}],"matriz_cumplimiento":[{"componente":"","estado":"acreditado|parcial|no_acreditado","sustento":""}],"datos_no_identificados":[]}
    extraction_system="""Actúa como extractor jurídico OSIPTEL. Separa hechos de conclusiones.

REGLA DE ORIGEN DE LA OBLIGACIÓN: si el acto es una Resolución TRASU, identifica el mandato y todas sus obligaciones EXCLUSIVAMENTE en el campo parte_resolutiva_trasu suministrado. Debe provenir del artículo o numeral que declara FUNDADO el reclamo y ordena las medidas de cumplimiento. Está prohibido construir, completar o modificar la obligación usando antecedentes, considerandos, alegaciones de la empresa, cartas posteriores o pruebas de ejecución. Copia el mandato con fidelidad y luego sepáralo por componentes, periodos, montos y condiciones. Si la parte resolutiva no permite identificarlo, registra el dato como no identificado.

REGLA OBLIGATORIA — VARIOS NUMERALES EN LA PARTE RESOLUTIVA: la parte resolutiva de una Resolución TRASU suele contener más de un numeral. Identifica como obligación principal EXCLUSIVAMENTE el o los numerales que ordenan la acción sustantiva vinculada a la prestación del servicio en discusión (por ejemplo: reactivar, reponer, migrar, ajustar, devolver, activar, dar de baja, entre otras similares), junto con su propio plazo expreso. IGNORA por completo, y no los registres ni como obligación principal ni como obligaciones a evaluar, los numerales que ordenan: (a) comunicar o informar al usuario sobre las acciones de cumplimiento ya ejecutadas; (b) informar o acreditar ante el TRASU el cumplimiento; o (c) cualquier otro plazo posterior de reporte o trámite formal (p. ej. cinco o diez días hábiles para informar). Esas son infracciones distintas que esta herramienta no evalúa; no las menciones como incumplimiento ni las mezcles con la obligación principal.

Para cada prueba distingue ejecución efectiva de solicitud, programación, caso abierto o gestión en curso. Una afirmación de la empresa no prueba por sí sola el hecho. Cita el documento que respalda cada dato.

REGLA OBLIGATORIA para obligaciones de informar, brindar, remitir, entregar o trasladar información al usuario (aplica en especial a devoluciones en efectivo y comunicaciones equivalentes): una carta o correo simple no acredita por sí solo que el usuario recibió la información. Exige acuse, confirmación de recepción o entrega, cargo de notificación con fecha o equivalente. Si solo existe envío sin recepción acreditada, marca el componente como no_acreditado.

EXCEPCIÓN OBLIGATORIA — AJUSTES, ANULACIONES O DESCUENTOS EN LA FACTURACIÓN: cuando la obligación consiste en ajustar, anular o descontar un importe en la facturación (no en devolver dinero en efectivo), la ejecución se acredita con la captura de pantalla del sistema o el histórico del estado de cuenta que muestre que el ajuste coincide con el importe ordenado por el TRASU, conforme a los criterios de la materia "Facturación y cobro". En estos casos NO se exige acreditar que el usuario recibió una notificación o carta sobre el ajuste; dicha comunicación, si existe, es evidencia adicional pero no condición de cumplimiento. No confundas la obligación de ajustar la facturación con una obligación de informar al usuario.

No evalúes todavía PAS ni subsanación. Devuelve JSON conforme al esquema."""
    extraction=parse_json_response(gemini_text(extraction_system,json.dumps({"esquema":extraction_schema,"caso":payload},ensure_ascii=False),json_mode=True)) or {}
    if not isinstance(extraction,dict): raise ValueError("Gemini no devolvió una extracción jurídica válida")
    schema={"ficha":{"expediente":"","empresa_operadora":"","usuario_abonado":"","servicio":"","tipo_acto":"","numero_acto":"","fecha_notificacion_emision":"","plazo_cumplimiento":"","fecha_vencimiento":"","obligacion_principal":"","medios_probatorios":[]},"trazabilidad":[],"evaluacion_juridica":{"checklist":[],"sustento_breve":"","tipo_incumplimiento":""},"resultado":"Cumplió|Incumplió|Inejecutable","subsanacion_voluntaria":"aplica|no aplica|no corresponde","clasificacion":"PAS|NO PAS","parrafo_final":"","datos_pendientes":[]}
    system="""Eres analista jurídico senior de OSIPTEL. Usa SOLO la evidencia y fuentes entregadas. No inventes fechas, pruebas ni conclusiones. Lo faltante es 'No identificado' o 'Pendiente de verificación'.

JERARQUÍA DOCUMENTAL OBLIGATORIA:
1. Sigue literalmente instrucciones_juridicas.txt para determinar el tipo de acto, el orden del análisis, la prueba exigible y los datos que no pueden inferirse.
2. Aplica íntegramente criterios_evaluacion_obligatorios.txt como una INSTRUCCIÓN vinculante. Sus filas proceden literalmente de CRITERIOS DE EVALUACION DE CUMPLIMIENTO.xlsx; no son simples referencias ni ejemplos. Debes conservar el significado de PAS, NO PAS, PLAZO e INEJECUTABLE indicado en cada fila.
3. Aplica como reglas vinculantes las filas seleccionadas de PAUTAS PAS.xlsx, separando cumplimiento, razonabilidad, cese y subsanación voluntaria.
4. Redacta con la estructura y lenguaje de PLANTILLAS cumplimiento.docx cuando sea una Resolución TRASU.
5. Redacta con la estructura y lenguaje de PLANTILLAS DENUNCIAS ACTUALIZADAS.docx para cartas, denuncias, SAR, SARA o SAP.
6. Antes de concluir, identifica en el checklist el nombre del archivo y la fila o sección aplicada. Si no puedes identificar la regla utilizada, no emitas evaluación final y registra el dato como pendiente.
7. Está prohibido emitir una conclusión basada solamente en conocimiento general del modelo cuando contradiga cualquiera de esos documentos.

MÉTODO Y CONTROLES JURÍDICOS OBLIGATORIOS (en este orden):
0. Determina la materia y cita en el checklist las filas concretas de criterios y pautas usadas. No uses una regla de otra materia.
1. Evalúa cada componente y cada periodo del mandato. El cumplimiento parcial NO equivale a cumplimiento íntegro.
2. Una solicitud, programación, ticket, caso abierto, gestión en curso o comunicación interna sobre una ejecución futura NO acredita ejecución efectiva. Exige nota de crédito, recibo ajustado, histórico aplicado u otra constancia objetiva.
3. Si falta prueba de uno o más meses ordenados, concluye que no se acreditó la ejecución íntegra.
4. La subsanación voluntaria exige conjuntamente cese TOTAL de la conducta y reversión INTEGRAL de todos sus efectos antes del procedimiento. Que la materia no restrinja el servicio, por sí solo, jamás basta.
5. Si siguen periodos, montos o componentes pendientes, establece 'no aplica' para el eximente y explica la ausencia de cese total/reversión integral.
6. Usa exclusivamente la fecha de vencimiento calculada y no la recalcules.
7. Aplica las instrucciones, criterios y pautas PAS entregados, incluso si contradicen una inferencia general del modelo. La plantilla determina la estructura de redacción.
8. Contrasta obligación por obligación con la matriz de pruebas; no generalices el resultado de un componente a los demás.
9. Mantén separadas tres decisiones: (a) cumplimiento material dentro del plazo, (b) razonabilidad para una ejecución tardía y (c) eximente de subsanación voluntaria. No conviertas automáticamente una ejecución tardía en subsanación.
10. En descuentos por varios meses, las notas de crédito o recibos solo acreditan los periodos que identifican. Para meses futuros exige un medio que muestre el REGISTRO Y ACTIVACIÓN efectivos de esos descuentos y su periodo. Una carta que diga que se gestionó o programó no los acredita.
11. La etiqueta NO PAS de una fila solo procede si todos los hechos descritos en esa fila están acreditados. Si faltan meses, extremos o prueba objetiva, aplica la fila específica de incumplimiento/PAS.
12. Antes de redactar, construye el checklist con: mandato; plazo; prueba exigible; prueba existente; periodos acreditados; periodos pendientes; regla aplicada; conclusión. Si existe contradicción, prevalece la evidencia documental y la regla específica.
13. El párrafo final debe indicar obligatoriamente y de forma expresa: fecha de notificación, número y tipo de días del plazo verificado, y fecha de vencimiento. Copia esos tres datos literalmente de la ficha; está prohibido recalcularlos u omitir el plazo.
14. Cada conclusión debe derivarse de hechos mencionados inmediatamente antes. No concluyas cumplimiento, incumplimiento, cese, reversión o subsanación si la matriz no identifica la prueba y los periodos que sustentan esa conclusión.
15. El párrafo final debe seguir literalmente la redacción y estructura de frases de la plantilla aplicable (PLANTILLAS cumplimiento.docx o PLANTILLAS DENUNCIAS ACTUALIZADAS.docx). Está prohibido agregar datos que la plantilla no contempla en esa oración, como el número de la resolución, carta, SARA, SAR, SAP o resolución de primera instancia, salvo que la plantilla lo incluya expresamente en su texto.

Devuelve JSON válido conforme al esquema y un párrafo final completo, cronológico, con obligación, pruebas por periodo, contraste, conclusión y análisis separado de subsanación."""
    user=json.dumps({"esquema":schema,"caso":{k:v for k,v in payload.items() if k!="documentos"},"extraccion_probatoria":extraction,"fuentes":sources},ensure_ascii=False)
    result=parse_json_response(gemini_text(system,user,json_mode=True)) or {}
    if not isinstance(result,dict): raise ValueError("Gemini no devolvió una evaluación jurídica válida")
    if not isinstance(result.get("ficha"),dict): result["ficha"]={}
    if not isinstance(result.get("evaluacion_juridica"),dict): result["evaluacion_juridica"]={}
    # Verified spreadsheet values are authoritative; the model cannot recalculate them.
    result.setdefault("ficha",{})["expediente"]=payload.get("expediente_detectado","No identificado")
    result["ficha"]["fecha_notificacion_emision"]=payload.get("fecha_verificada","No identificado")
    result["ficha"]["plazo_cumplimiento"]=payload.get("plazo_verificado","Pendiente de verificación")
    result["ficha"]["fecha_vencimiento"]=payload.get("fecha_vencimiento","Pendiente de verificación")
    resolutive_obligation=extraction.get("obligacion_extraida_parte_resolutiva") or {}
    if isinstance(resolutive_obligation,dict) and str(resolutive_obligation.get("texto","")).strip():
        result["ficha"]["obligacion_principal"]=str(resolutive_obligation["texto"]).strip()
    elif payload.get("tipo_acto")=="Resolución TRASU":
        raise ValueError("La parte resolutiva fue localizada, pero no se pudo extraer el mandato que declara fundado el reclamo")
    # Deterministic legal guard: pending/programmed work is not full execution.
    statuses={str(x.get("estado","")).lower() for x in (extraction.get("medios_probatorios") or []) if isinstance(x,dict)}
    matrix={str(x.get("estado","")).lower() for x in (extraction.get("matriz_cumplimiento") or []) if isinstance(x,dict)}
    # A single unaccredited piece of evidence (e.g. an incidental notification letter) must not
    # override compliance by itself; only the per-obligation matrix (which the main evaluation
    # step builds while reading criterios_evaluacion_obligatorios.txt) decides that. "programado"/
    # "en_curso" are the exception: they mean the action was never finished, regardless of which
    # document reports it.
    incomplete=bool(statuses & {"programado","en_curso"} or matrix & {"parcial","no_acreditado"})
    if incomplete:
        result["resultado"]="Incumplió"
        result["subsanacion_voluntaria"]="no aplica"
        result["clasificacion"]="PAS"
        result.setdefault("evaluacion_juridica",{}).setdefault("checklist",[]).append(
            "Regla obligatoria: existen periodos programados, en curso o no acreditados; no hubo ejecución íntegra, cese total ni reversión integral.")
        locked={
            "ficha":result.get("ficha",{}),"extraccion_probatoria":extraction,
            "resultado_obligatorio":"Incumplió","subsanacion_obligatoria":"no aplica",
            "clasificacion_obligatoria":"PAS","fuentes_aplicables":sources,
        }
        rewrite_system="""Redacta un único párrafo jurídico conforme a la plantilla TRASU. Las conclusiones indicadas como obligatorias están bloqueadas y no puedes modificarlas. Debes indicar literalmente la fecha de notificación, el plazo de cumplimiento (número y tipo de días) y la fecha de vencimiento que aparecen en la ficha; no puedes recalcularlos ni omitirlos. No afirmes que una programación, solicitud, carta o gestión en curso acredita ejecución. Explica que, al quedar periodos sin prueba de registro y activación efectiva, no hubo ejecución íntegra, cese total ni reversión integral. No apliques el eximente por el solo hecho de que no exista restricción del servicio. Devuelve JSON {parrafo_final:string}."""
        rewritten=parse_json_response(gemini_text(rewrite_system,json.dumps(locked,ensure_ascii=False),json_mode=True)) or {}
        if not isinstance(rewritten,dict): rewritten={}
        result["parrafo_final"]=rewritten.get("parrafo_final",result.get("parrafo_final",""))
    # A response without a paragraph is never a completed evaluation.
    if not str(result.get("parrafo_final","")).strip():
        locked={"ficha":result.get("ficha",{}),"extraccion_probatoria":extraction,
                "resultado":result.get("resultado"),"subsanacion_voluntaria":result.get("subsanacion_voluntaria"),
                "clasificacion":result.get("clasificacion"),"fuentes_aplicables":sources}
        fallback_system="""Redacta el párrafo jurídico final conforme a la plantilla aplicable. Relaciona en orden: acto y notificación; plazo y vencimiento de la ficha; mandato; alegaciones; pruebas objetivas; contraste por cada obligación o periodo; conclusión; y subsanación. No cambies el resultado, la clasificación ni la subsanación ya determinadas. No inventes. Devuelve JSON {parrafo_final:string}."""
        try:
            fallback=parse_json_response(gemini_text(fallback_system,json.dumps(locked,ensure_ascii=False),json_mode=True)) or {}
            if isinstance(fallback,dict): result["parrafo_final"]=str(fallback.get("parrafo_final","")).strip()
        except Exception:
            result["parrafo_final"]=""
    if not str(result.get("parrafo_final","")).strip():
        result["parrafo_final"]=deterministic_paragraph(result,extraction)
    paragraph=str(result.get("parrafo_final","")).strip()
    # The templates never cite the act's identifying number; strip it if the model added it anyway.
    numero_acto=str(result.get("ficha",{}).get("numero_acto","")).strip()
    if numero_acto and numero_acto.lower() not in {"no identificado","pendiente de verificación","pendiente de verificacion"}:
        paragraph=re.sub(r"\s*N\.?[°º]\s*"+re.escape(numero_acto),"",paragraph)
    # Avoid an accidental exact duplication of the generated paragraph.
    half=len(paragraph)//2
    if len(paragraph)%2==0 and paragraph[:half]==paragraph[half:]:
        paragraph=paragraph[:half].strip()
    result["parrafo_final"]=paragraph
    return result

def regenerate_paragraph(result: dict[str,Any]) -> str:
    context={"ficha":result.get("ficha",{}),"evaluacion_juridica":result.get("evaluacion_juridica",{}),"resultado":result.get("resultado"),"subsanacion_voluntaria":result.get("subsanacion_voluntaria"),"clasificacion":result.get("clasificacion"),"datos_pendientes":result.get("datos_pendientes",[]),"instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore")[:40000]}
    response=gemini_text("Regenera únicamente el párrafo jurídico final con los datos aportados. No inventes ni completes faltantes. Devuelve JSON {parrafo_final:string}.",json.dumps(context,ensure_ascii=False),json_mode=True)
    return parse_json_response(response)["parrafo_final"]

def to_row(result: dict, documents: list[str]) -> dict:
    f=result.get("ficha",{}); e=result.get("evaluacion_juridica",{})
    return dict(zip(COLUMNAS,[f.get("expediente",""),f.get("empresa_operadora",""),f.get("usuario_abonado",""),f.get("servicio",""),f.get("tipo_acto",""),f.get("numero_acto",""),f.get("fecha_notificacion_emision",""),f.get("fecha_vencimiento",""),f.get("obligacion_principal",""),result.get("resultado",""),e.get("tipo_incumplimiento",""),result.get("subsanacion_voluntaria",""),result.get("clasificacion",""),e.get("sustento_breve",""),result.get("parrafo_final",""),"; ".join(map(str,result.get("datos_pendientes",[]))),"; ".join(documents),datetime.now().strftime("%Y-%m-%d %H:%M")]))

def excel_bytes(row: dict | None=None) -> bytes:
    existing=pd.read_excel(HISTORIAL,dtype=str) if HISTORIAL.exists() else pd.DataFrame(columns=COLUMNAS)
    if row is not None: existing=pd.concat([existing,pd.DataFrame([row])],ignore_index=True)
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w: existing.to_excel(w,index=False,sheet_name="Evaluaciones")
    return b.getvalue()

for k,v in {"result":None,"docs":[],"texts":{},"notice":None,"due":None,
            "analysis_error":None,"analysis_status":None,"upload_signature":None,
            "debug_resolutivo":None}.items():
    st.session_state.setdefault(k,v)

left,center,right=st.columns([1.05,2.25,1.05],gap="large")
with left:
    st.markdown('<div class="light">Expediente particular</div>',unsafe_allow_html=True)
    uploads=st.file_uploader("Cargar expediente",accept_multiple_files=True,type=["pdf","docx","xlsx","xls","csv","txt","png","jpg","jpeg","tif","tiff","zip","rar","7z"])
    if uploads:
        signature="|".join(f"{u.name}:{hashlib.sha256(u.getvalue()).hexdigest()}" for u in uploads)
        if signature!=st.session_state.upload_signature:
            st.session_state.result=None; st.session_state.analysis_error=None
            st.session_state.analysis_status="Archivos nuevos listos para analizar"
            st.session_state.upload_signature=signature
        st.session_state.docs=[u.name for u in uploads]
        for u in uploads:
            kind=classify(u.name,""); st.caption(f"✓ {u.name} · {kind}")
    ok_sources,missing=fuente_status()
    st.divider(); st.markdown('<div class="light">Fuentes permanentes</div>',unsafe_allow_html=True)
    st.success("Fuentes + instrucciones jurídicas y criterios obligatorios disponibles") if ok_sources else st.error(f"Faltan {len(missing)} fuentes o instrucciones")
    if missing:
        with st.expander("Ver faltantes"): st.write("\n".join(f"• {x}" for x in missing))
    analyze=st.button("✦ Analizar expediente con IA",type="primary",use_container_width=True,disabled=not uploads)

if analyze:
    st.session_state.result=None; st.session_state.analysis_error=None
    st.session_state.analysis_status="Análisis iniciado"
    if not ok_sources: st.error("La herramienta no puede emitir evaluación final porque no tiene acceso a las fuentes permanentes")
    elif not get_api_key(): st.error("Configure GEMINI_API_KEY en .env o en los secretos de Streamlit.")
    else:
        with st.spinner("Procesando y contrastando el expediente..."):
            case_dir=Path(tempfile.mkdtemp(prefix="caso_",dir=TEMP))
            try:
                paths=[]
                for u in uploads:
                    p=case_dir/safe_name(u.name); p.write_bytes(u.getbuffer()); paths.append(p)
                paths+=extract_archives(case_dir)
                unique_paths=[]; seen_hashes=set()
                for p in paths:
                    if p.suffix.lower() in {".zip",".rar",".7z"}: continue
                    try: digest=hashlib.sha256(p.read_bytes()).hexdigest()
                    except Exception: digest=str(p.resolve())
                    if digest not in seen_hashes:
                        seen_hashes.add(digest); unique_paths.append(p)
                paths=unique_paths
                texts={p.name:read_file(p) for p in paths if p.suffix.lower() not in {".zip",".rar",".7z"}}
                if not texts or not any(str(x).strip() for x in texts.values()):
                    raise ValueError("No se pudo extraer texto de los archivos del expediente.")
                st.session_state.analysis_status="Documentos leídos; verificando expediente y plazo"
                combined="\n\n".join(f"### {n}\n{t}" for n,t in texts.items())
                document_types=[classify(n,t) for n,t in texts.items()]; tipo=next((x for x in document_types if x!="No identificado"),"No identificado")
                resolutive=extract_resolutive_part(texts) if tipo=="Resolución TRASU" else None
                st.session_state["debug_resolutivo"]=resolutive
                if tipo=="Resolución TRASU" and not resolutive:
                    raise ValueError("No se identificó la parte resolutiva que declara fundado el reclamo; no es posible establecer la obligación sin esa sección")
                searchable="\n".join(u.name for u in uploads)
                expediente=parse_trasu_name(searchable) or identify_exact_expediente(searchable) or "No identificado"
                notice=None; due=None; term=None
                if tipo=="Resolución TRASU":
                    notice=exact_notification(expediente) if expediente!="No identificado" else None
                    if not notice:
                        st.session_state.analysis_error="No se encontró coincidencia exacta del expediente en notificaciones mayo.xlsx. Expediente detectado: "+expediente
                    else:
                        deadline=calculate_due(notice,combined)
                        if deadline: due,term=deadline
                        if not deadline: st.session_state.analysis_error="No se pudo determinar el plazo o calcular el vencimiento con CONTADOR DE PLAZOS - TRASU 2026.xlsx"
                if tipo!="Resolución TRASU" or (notice and due):
                    st.session_state.analysis_status="Aplicando instrucciones, criterios, pautas y plantillas"
                    payload={"tipo_acto":tipo,"expediente_detectado":expediente,"fecha_verificada":notice or "No identificado","plazo_verificado":term or "Pendiente de verificación","fecha_vencimiento":due or "Pendiente de verificación","parte_resolutiva_trasu":resolutive or "No corresponde","documentos":texts}
                    st.session_state.result=ai_evaluate(payload); st.session_state.texts=texts
                    st.session_state.analysis_status="Evaluación completada"
            except Exception as e:
                st.session_state.analysis_error=f"No se pudo completar el análisis: {type(e).__name__}: {e}"
            finally: shutil.rmtree(case_dir,ignore_errors=True)

if st.session_state.analysis_error:
    st.error(st.session_state.analysis_error)
if st.session_state.debug_resolutivo:
    with st.expander("Ver texto de la parte resolutiva detectada (para diagnóstico)"):
        st.text(st.session_state.debug_resolutivo)
elif st.session_state.analysis_status:
    st.info(st.session_state.analysis_status)

r=st.session_state.result
with center:
    st.markdown('<div class="light">Ficha editable del expediente</div>',unsafe_allow_html=True)
    f=(r or {}).get("ficha",{})
    a,b=st.columns(2)
    expediente=a.text_input("Expediente",f.get("expediente","No identificado")); empresa=b.text_input("Empresa operadora",f.get("empresa_operadora","No identificado"))
    usuario=a.text_input("Usuario o abonado",f.get("usuario_abonado","No identificado")); servicio=b.text_input("Servicio",f.get("servicio","No identificado"))
    tipo=a.text_input("Tipo de acto",f.get("tipo_acto","No identificado")); numero=b.text_input("Número de resolución o carta",f.get("numero_acto","No identificado"))
    notif=a.text_input("Fecha de notificación o emisión",f.get("fecha_notificacion_emision","No identificado")); plazo=b.text_input("Plazo de cumplimiento",f.get("plazo_cumplimiento","Pendiente de verificación"))
    vence=st.text_input("Fecha máxima de vencimiento",f.get("fecha_vencimiento","Pendiente de verificación"))
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

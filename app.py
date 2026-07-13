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
APP_VERSION = "2026.07.13-17"
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
st.caption(f"Versión {APP_VERSION}")

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
        # Do not repeat an identical rejected request: repeated RECITATION calls
        # consume the small free quota without improving the document extraction.
        for attempt in range(1):
            try:
                response=client.models.generate_content(model=model,contents=user,config=config)
                candidates=response.candidates or []
                finish=str(candidates[0].finish_reason) if candidates else "SIN_CANDIDATOS"
                text=response.text if candidates and candidates[0].content and candidates[0].content.parts else None
                if not text: raise RuntimeError(f"El modelo {model} no devolvió contenido utilizable (motivo: {finish}).")
                return text
            except Exception as e:
                errors.append(f"{model} (intento {attempt+1}): {e}")
                if "RECITATION" not in str(e): break
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
            reader=PdfReader(str(path))
            text="\n".join((p.extract_text() or "") for p in reader.pages)
            # Hybrid PDFs often contain a short text layer plus printers,
            # screenshots or scanned evidence as images. A text-length check
            # alone silently omitted those facts. OCR every page locally when
            # raster images are present and append it to the extracted text.
            try:
                has_raster_images=any(bool(getattr(page,"images",[])) for page in reader.pages)
            except Exception:
                has_raster_images=False
            if len(text.strip())<80 or has_raster_images:
                try:
                    import pytesseract
                    from pdf2image import convert_from_path
                    images=convert_from_path(str(path),dpi=170,thread_count=2)
                    ocr_text="\n".join(
                        f"--- OCR PÁGINA {i+1} ---\n"+pytesseract.image_to_string(image,lang="spa")
                        for i,image in enumerate(images))
                    if ocr_text.strip(): text=(text+"\n\n"+ocr_text).strip()
                except Exception: pass
            if len(text.strip())<80:
                return vision_ocr(path)
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
    """OCR scanned images in small batches, prioritizing the end of legal resolutions."""
    schema={"tipo_documento":"","expediente":"","numero_acto":"","fecha_acto":"",
            "ha_resuelto":[{"numeral":"","obligacion_o_disposicion":"","plazo":"",
                             "destinatario":"","es_obligacion_accesoria":False}],
            "hechos_probatorios_relevantes":[]}
    instruction="""Analiza las imágenes de este documento jurídico y extrae hechos estructurados; NO transcribas íntegramente el documento ni reproduzcas párrafos largos.
Identifica tipo de documento, expediente, número y fecha del acto. Si aparece HA RESUELTO, extrae por separado cada numeral, describiendo fielmente pero de forma concisa su obligación o disposición, su plazo, destinatario y si es una obligación accesoria de informar posteriormente el cumplimiento al usuario o al TRASU. No inventes datos ilegibles.
Para documentos probatorios, resume únicamente fechas, importes, acciones efectivamente ejecutadas y constancias visibles. Devuelve JSON conforme al esquema."""
    def facts_text(raw: str) -> str:
        data=parse_json_response(raw) if raw else {}
        if not isinstance(data,dict): return ""
        lines=[f"TIPO DE DOCUMENTO: {data.get('tipo_documento','')}",
               f"EXPEDIENTE: {data.get('expediente','')}",f"NÚMERO DE ACTO: {data.get('numero_acto','')}",
               f"FECHA DEL ACTO: {data.get('fecha_acto','')}"]
        operative=data.get("ha_resuelto") or []
        if operative:
            lines.append("HA RESUELTO")
            for item in operative:
                if not isinstance(item,dict): continue
                lines.append(f"NUMERAL {item.get('numeral','')}: {item.get('obligacion_o_disposicion','')} | PLAZO: {item.get('plazo','')} | DESTINATARIO: {item.get('destinatario','')} | ACCESORIA: {item.get('es_obligacion_accesoria',False)}")
        for fact in (data.get("hechos_probatorios_relevantes") or []): lines.append(f"HECHO PROBATORIO: {fact}")
        return "\n".join(lines)
    try:
        from PIL import Image
        Image.MAX_IMAGE_PIXELS=None
        source=Image.open(path)
        total=max(1,int(getattr(source,"n_frames",1)))
        # HA RESUELTO is normally on the final pages. For very large TIFF files,
        # read the cover and the final 18 pages instead of silently stopping at
        # the first 20 pages, which previously omitted the operative section.
        if total<=30:
            indices=list(range(total))
        else:
            indices=sorted(set(list(range(min(6,total)))+list(range(max(0,total-18),total))))
        batches=[indices[i:i+5] for i in range(0,len(indices),5)]
        transcriptions=[]; errors=[]
        for batch in batches:
            content=[instruction+"\nESQUEMA: "+json.dumps(schema,ensure_ascii=False)]
            for index in batch:
                source.seek(index)
                image=source.convert("RGB"); image.thumbnail((1500,1500))
                buf=io.BytesIO(); image.save(buf,format="JPEG",quality=78,optimize=True)
                content.append(f"--- Página {index+1} de {total} ---")
                content.append(types.Part.from_bytes(data=buf.getvalue(),mime_type="image/jpeg"))
            try:
                text=gemini_text("Eres un extractor de hechos jurídicos visibles en imágenes. Evita transcripciones extensas.",content,json_mode=True)
                converted=facts_text(text)
                if converted.strip(): transcriptions.append(converted)
            except Exception as batch_error:
                errors.append(f"páginas {batch[0]+1}-{batch[-1]+1}: {batch_error}")
        combined_transcription="\n\n".join(transcriptions)
        # Recovery used only when the ordinary OCR read a TRASU resolution but
        # omitted its operative section.  Inspect the final pages individually,
        # in reverse order, and stop as soon as HA RESUELTO is recovered.  This
        # supplements the first extraction; it does not replace any other data.
        looks_like_trasu=("trasu" in path.name.lower() or
                          "resoluci" in path.name.lower() or
                          "TRASU" in combined_transcription.upper())
        if looks_like_trasu and "HA RESUELTO" not in combined_transcription.upper():
            operative_schema={"ha_resuelto":[{
                "numeral":"", "obligacion_o_disposicion":"", "plazo":"",
                "destinatario":"", "es_obligacion_accesoria":False
            }]}
            operative_instruction="""Examina solamente esta pagina de una resolucion TRASU.
Si contiene el titulo HA RESUELTO, SE RESUELVE o la continuacion de sus numerales, resume TODOS los numerales visibles en ha_resuelto. Conserva la relacion entre el mandato material y su plazo aunque aparezcan en numerales consecutivos. Incluye tambien las obligaciones accesorias y marcalas como tales. No transcribas parrafos completos y no inventes texto. Si esta pagina no contiene parte resolutiva, devuelve ha_resuelto como lista vacia. Devuelve solo JSON conforme al esquema."""
            recovery_errors=[]
            # One recovery request at most. Send the central body of the final
            # pages together instead of making one or two API calls per page.
            # This keeps the rule general while protecting the user's quota.
            try:
                recovery_content=[operative_instruction+"\nESQUEMA: "+json.dumps(operative_schema,ensure_ascii=False)]
                for index in indices[-6:]:
                    source.seek(index)
                    original=source.convert("RGB")
                    width,height=original.size
                    body=original.crop((0,int(height*.12),width,int(height*.92))) if height>800 else original
                    body.thumbnail((1500,1500))
                    buf=io.BytesIO(); body.save(buf,format="JPEG",quality=80,optimize=True)
                    recovery_content.append(f"Pagina {index+1} de {total}, cuerpo central")
                    recovery_content.append(types.Part.from_bytes(data=buf.getvalue(),mime_type="image/jpeg"))
                # The previous working version recovered this section as plain
                # text. JSON extraction may legally succeed while returning an
                # empty ha_resuelto list, so use one concise plain-text request
                # here and normalize its numbered output under the heading.
                recovery_content[0]="""Localiza en estas paginas la parte resolutiva de la Resolucion TRASU.
Devuelve exclusivamente un resumen fiel con este formato:
HA RESUELTO
NUMERAL 1: disposicion y plazo visible
NUMERAL 2: disposicion y plazo visible
Incluye todos los numerales resolutivos visibles y conserva las referencias entre numerales consecutivos. No agregues antecedentes, considerandos ni hechos de otros documentos. Si no aparece ninguna parte resolutiva, responde exactamente NO_ENCONTRADO."""
                raw=gemini_text("Lee visualmente la parte resolutiva y resume sus numerales sin copiar parrafos extensos.",recovery_content,json_mode=False)
                normalized_raw=str(raw or "").strip()
                if normalized_raw and "NO_ENCONTRADO" not in normalized_raw.upper() and re.search(r"\bNUMERAL\s+\w+",normalized_raw,re.I):
                    if "HA RESUELTO" not in normalized_raw.upper():
                        normalized_raw="HA RESUELTO\n"+normalized_raw
                    combined_transcription=(combined_transcription+"\n\n"+normalized_raw).strip()
            except Exception as recovery_error:
                recovery_errors.append(f"recuperacion conjunta: {recovery_error}")
        if combined_transcription:
            return combined_transcription
        # Second route: let the Files API ingest the original TIFF. This avoids
        # Pillow/libtiff decoder limitations that affect some official scans.
        try:
            client=gemini_client()
            uploaded=client.files.upload(file=str(path))
            prompt=instruction+"\nESQUEMA: "+json.dumps(schema,ensure_ascii=False)
            raw=gemini_text("Eres un extractor de hechos jurídicos. Evita transcripciones extensas.",[uploaded,prompt],json_mode=True)
            converted=facts_text(raw)
            if converted.strip(): return converted
            errors.append("el archivo TIFF original no produjo texto")
        except Exception as upload_error:
            errors.append(f"lectura directa del TIFF: {upload_error}")
        raise RuntimeError("; ".join(errors) if errors else "ningún método produjo texto")
    except Exception as e:
        # If page decoding itself failed, still try the original file once.
        try:
            client=gemini_client()
            uploaded=client.files.upload(file=str(path))
            prompt=instruction+"\nESQUEMA: "+json.dumps(schema,ensure_ascii=False)
            raw=gemini_text("Eres un extractor de hechos jurídicos. Evita transcripciones extensas.",[uploaded,prompt],json_mode=True)
            converted=facts_text(raw)
            if converted.strip(): return converted
        except Exception as upload_error:
            return f"[OCR no disponible para {path.name}: conversión por páginas: {e}; lectura directa: {upload_error}]"
        return f"[OCR no disponible para {path.name}: {e}]"

def classify(name: str, text: str) -> str:
    # TRASU must win when it is present in the filename or document text.
    # SAR/SARA/SAP are acronyms: substring matching wrongly classified ordinary
    # words containing "sar" and caused valid TRASU resolutions to be skipped.
    s=name+" "+text[:6000]
    folded=unicodedata.normalize("NFD",s.upper())
    folded="".join(c for c in folded if unicodedata.category(c)!="Mn")
    if re.search(r"(?<![A-Z0-9])TRASU(?![A-Z0-9])",folded) or "TRASU" in name.upper():
        return "Resolución TRASU"
    for acronym,label in (("SARA","SARA"),("SAR","SAR"),("SAP","SAP")):
        if re.search(rf"(?<![A-Z0-9]){acronym}(?![A-Z0-9])",folded): return label
    low=folded.lower()
    for key,label in (("denuncia","Denuncia"),("carta","Carta"),("primera instancia","Resolución de primera instancia")):
        if key in low: return label
    return "No identificado"

def extract_resolutive_part(documents: dict[str,str]) -> str | None:
    """Extract the operative section; obligations may not be inferred elsewhere."""
    candidates=[]
    for name,text in documents.items():
        if classify(name,text)!="Resolución TRASU": continue
        upper=unicodedata.normalize("NFD",text.upper())
        upper="".join(c for c in upper if unicodedata.category(c)!="Mn")
        upper_name=unicodedata.normalize("NFD",name.upper())
        upper_name="".join(c for c in upper_name if unicodedata.category(c)!="Mn")
        resolution_document=(
            bool(re.search(r"RESOLUCION[^\n]{0,40}FINAL",upper_name)) or
            bool(re.search(r"TRIBUNAL\s+ADMINISTRATIVO[\s\S]{0,500}RESOLUCION\s+FINAL",upper))
        )
        # A letter or an evidence annex may mention TRASU, but that does not
        # turn it into the resolution whose operative section defines the duty.
        if not resolution_document: continue
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

def extract_trasu_mandate_and_term(resolutive: str) -> dict[str,Any]:
    """Read HA RESUELTO as a whole and associate the principal mandate with its own term."""
    schema={
        "obligacion_principal":"",
        "plazo_principal_textual":"",
        "numero_plazo":None,
        "tipo_plazo":"dias_habiles|dias_calendario|meses|no_identificado",
        "numerales_fuente":[],
        "obligaciones_accesorias_excluidas":[],
    }
    system="""Lee íntegramente la sección HA RESUELTO de una Resolución TRASU.
Identifica la obligación material principal relacionada con la prestación o controversia del servicio y el plazo expresamente asociado a SU ejecución. Lee conjuntamente todos los numerales: la obligación y su plazo pueden estar en puntos consecutivos.
No uses una lista cerrada de verbos y no exijas que el mandato esté en el mismo numeral que declara fundado el reclamo.
Excluye las obligaciones posteriores de informar el cumplimiento al usuario, informar al TRASU, remitir constancias, acreditar comunicaciones o reportar acciones, junto con los plazos de esas obligaciones accesorias. Nunca confundas esos plazos con el plazo principal.
Los considerandos solo pueden aclarar una referencia del mandato resolutivo; no pueden crear ni ampliar la obligación.
Copia literalmente el plazo principal. Si no está legible, usa no_identificado; no adivines ni apliques un plazo general.
Devuelve JSON válido conforme al esquema."""
    raw=gemini_text(system,json.dumps({"esquema":schema,"ha_resuelto":resolutive},ensure_ascii=False),json_mode=True)
    data=parse_json_response(raw) if raw else {}
    if not isinstance(data,dict):
        raise ValueError("No se pudo interpretar la sección HA RESUELTO")
    if not str(data.get("obligacion_principal") or "").strip():
        raise ValueError("No se pudo identificar la obligación material principal en HA RESUELTO")
    if not str(data.get("plazo_principal_textual") or "").strip() or str(data.get("tipo_plazo") or "") == "no_identificado":
        raise ValueError("No se pudo identificar el plazo vinculado a la obligación principal en HA RESUELTO")
    return data

def exact_notification(expediente: str) -> str | None:
    path=FUENTES/"notificaciones mayo.xlsx"
    for _,df in pd.read_excel(path,sheet_name=None,dtype=str).items():
        cols={str(c).strip().upper():c for c in df.columns}
        date_key="FEC_NOT_EMP_ELE_TEXTO" if "FEC_NOT_EMP_ELE_TEXTO" in cols else ("FEC_NOT_EMP_ELE" if "FEC_NOT_EMP_ELE" in cols else None)
        if "NRO_EXPEDIENTE" in cols and date_key:
            values=df[cols["NRO_EXPEDIENTE"]].fillna("").astype(str).str.strip()
            target=re.sub(r"[^A-Z0-9]","",expediente.upper())
            normalized=values.str.upper().str.replace(r"[^A-Z0-9]","",regex=True)
            hit=df.loc[normalized==target,cols[date_key]]
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
    match=re.search(r"(\d{7})(20\d{2})TRASU(STRA|STRQJ)",compact)
    if not match: return None
    suffix="ST-RA" if match.group(3)=="STRA" else "ST-RQJ"
    return f"{match.group(1)}-{match.group(2)}/TRASU/{suffix}"

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
    normalized_case=unicodedata.normalize("NFD",case_context.lower())
    normalized_case="".join(c for c in normalized_case if unicodedata.category(c)!="Mn")
    discount_case=any(k in normalized_case for k in (
        "descuento","ajuste","facturacion","importe","nota de credito",
        "oferta","promocion","beneficio recurrente"))
    ranked=[]
    for sheet,df in book.items():
        for idx,row in df.fillna("").iterrows():
            cells=[str(x).strip() for x in row.tolist() if str(x).strip()]
            if not cells: continue
            line=" | ".join(cells)
            tokens=_legal_tokens(line)
            score=len(query & tokens)
            low=line.lower()
            # Only truly transversal rules are boosted for every matter.
            # Discount/programming rules are boosted exclusively when the case
            # itself concerns billing, adjustments, offers or promotions.
            general=any(k in low for k in (
                "cese total","reversión total","subsanación voluntaria",
                "análisis por cada mandato","conclusión es por resolución",
                "cumplimiento parcial"))
            discount_rule=any(k in low for k in (
                "registro y activación de los descuentos","descuento recurrente",
                "ajuste recurrente","meses pendientes","meses restantes"))
            general=general or (discount_case and discount_rule)
            if general: score+=100
            if score: ranked.append((score,sheet,int(idx)+1,line))
    ranked.sort(key=lambda x:(-x[0],x[1],x[2]))
    selected=ranked[:limit]
    return "\n".join(f"[{name} / {s} / fila {r}] {line}" for _,s,r,line in selected)

def _fold_legal_text(value: Any) -> str:
    text=unicodedata.normalize("NFD",str(value or "").lower())
    return "".join(c for c in text if unicodedata.category(c)!="Mn")

def _criteria_rows() -> list[dict[str,Any]]:
    """Parse the institutional instruction while preserving each row's matter."""
    raw=CRITERIOS_INSTRUCCION.read_text("utf-8",errors="ignore")
    matches=list(re.finditer(r"(?ms)^FILA\s+(\d+):\s*(.*?)(?=^FILA\s+\d+:|\Z)",raw))
    known_matters={
        "Facturación y cobro","Calidad e idoneidad",
        "Incumplimiento de condiciones contractuales, ofertas y promociones",
        "Corte o Baja Injustificada","Instalación, activación o traslado del servicio",
        "Falta de ejecución de baja o traslado del servicio","Recargas",
        "Contratación no solicitada","Migración","Portabilidad","Todas las materias",
        "Suspensión por uso prohibido","Activación de promoción",
        "Devoluciones en Centros de Atención de la empresa Operadora",
        "Portabilidad no autorizada","Cambio de número no autorizado",
    }
    folded_known={_fold_legal_text(x):x for x in known_matters}
    current=""
    rows=[]
    for match in matches:
        number=int(match.group(1))
        body=re.sub(r"\s+"," ",match.group(2)).strip()
        parts=[p.strip() for p in body.split("|")]
        first=_fold_legal_text(parts[0] if parts else "")
        if first in folded_known:
            current=folded_known[first]
        matter="Notificaciones" if number==66 else current
        if number>=3:
            rows.append({"fila":number,"materia":matter,"texto":body})
    return rows

def select_applicable_criteria(obligation: str) -> tuple[str,str,list[int]]:
    """Route the verified mandate to its own matrix block before legal reasoning."""
    rows=_criteria_rows()
    folded=_fold_legal_text(obligation)
    routing={
        "Corte o Baja Injustificada":("reconect","reactiv","restablec","corte injust","baja injust","servicio suspend"),
        "Calidad e idoneidad":("averia","calidad","idoneidad","reparar servicio","operatividad del servicio"),
        "Incumplimiento de condiciones contractuales, ofertas y promociones":("oferta","promocion","beneficio","descuento","condicion contractual"),
        "Facturación y cobro":("facturacion","cobro","ajustar importe","anular importe","nota de credito","estado de cuenta"),
        "Instalación, activación o traslado del servicio":("instalar","instalacion","activar el servicio","activacion del servicio","trasladar el servicio"),
        "Falta de ejecución de baja o traslado del servicio":("dar de baja","baja del servicio","ejecutar la baja","traslado del servicio"),
        "Contratación no solicitada":("contratacion no solicitada","servicio no solicitado"),
        "Migración":("migracion","migrar","cambio de plan"),
        "Portabilidad no autorizada":("portabilidad no autorizada",),
        "Portabilidad":("portabilidad","portar el numero"),
        "Cambio de número no autorizado":("cambio de numero no autorizado",),
        "Recargas":("recarga","saldo recargado"),
        "Suspensión por uso prohibido":("suspension por uso prohibido","rentseg","imei"),
        "Activación de promoción":("activar promocion","activacion de promocion"),
        "Devoluciones en Centros de Atención de la empresa Operadora":("devolucion en centro","recoger devolucion","centro de atencion","devolver importe","devolucion de importe"),
    }
    scores=[]
    for order,(matter,aliases) in enumerate(routing.items()):
        score=0
        for alias in aliases:
            if alias in folded:
                score+=max(2,len(alias.split()))
        matter_tokens=_legal_tokens(matter)
        score+=len(_legal_tokens(obligation)&matter_tokens)
        scores.append((score,-order,matter))
    scores.sort(reverse=True)
    primary=scores[0][2] if scores and scores[0][0]>0 else "Todas las materias"
    selected=[r for r in rows if r["materia"] in {primary,"Todas las materias","Notificaciones"}]
    ids=[int(r["fila"]) for r in selected]
    header=("Fuente consultada: criterios_evaluacion_obligatorios.txt. "
            f"Materia seleccionada desde la obligación principal: {primary}. "
            "Solo las filas siguientes pueden utilizarse para este razonamiento; "
            "las filas de otras materias quedan excluidas para evitar cruces.")
    text=header+"\n"+"\n".join(f"[FILA {r['fila']} / {r['materia']}] {r['texto']}" for r in selected)
    return primary,text,ids

def legal_sources(case_context: str, template: str, obligation: str) -> dict[str,str]:
    matter,criteria,criteria_ids=select_applicable_criteria(obligation)
    return {
        "instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore"),
        "materia_criterios_seleccionada":matter,
        "filas_criterios_seleccionadas":", ".join(map(str,criteria_ids)),
        "criterios_evaluacion_aplicables":criteria,
        "plantilla_aplicable":source_text(template,80000),
        "pautas_pas_aplicables":relevant_excel_rules("PAUTAS PAS.xlsx",case_context,26),
    }

def calculate_due(notification: str, context: str) -> tuple[str,str] | None:
    """Calculate only the principal term already extracted from HA RESUELTO."""
    try:
        normalized=unicodedata.normalize("NFD",context.lower())
        normalized="".join(c for c in normalized if unicodedata.category(c)!="Mn")
        start=parse_excel_date(notification)
        number_match=re.search(r"(?:\(|\b)(\d{1,3})\)?\s*(?:dias?|mes(?:es)?)",normalized,re.I)
        word_numbers={"un":1,"uno":1,"dos":2,"tres":3,"cuatro":4,"cinco":5,"seis":6,
                      "siete":7,"ocho":8,"nueve":9,"diez":10,"quince":15,"veinte":20}
        word_match=re.search(r"\b("+"|".join(word_numbers)+r")\b\s*(?:dias?|mes(?:es)?)",normalized,re.I)
        quantity=int(number_match.group(1)) if number_match else (word_numbers[word_match.group(1)] if word_match else None)
        if quantity is None:
            raise ValueError("el plazo principal no contiene una cantidad identificable")
        if "mes" in normalized:
            label=f"{quantity} mes"+("es" if quantity!=1 else "")
            return (start+pd.DateOffset(months=quantity)).strftime("%d/%m/%Y"),label
        if "calendario" in normalized or "naturales" in normalized:
            return (start.normalize()+pd.Timedelta(days=quantity)).strftime("%d/%m/%Y"),f"{quantity} días calendario"
        if "habil" not in normalized:
            raise ValueError("no se identificó si el plazo principal se computa en días hábiles")
        days=quantity
        term="un (1) día hábil" if days==1 else f"{days} días hábiles"
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

def normalize_legal_paragraph(paragraph: str, ficha: dict[str,Any], resultado: str="") -> str:
    """Enforce template wording and dd/mm/yyyy dates on every generation route."""
    text=str(paragraph or "").strip()
    months={"enero":"01","febrero":"02","marzo":"03","abril":"04","mayo":"05","junio":"06",
            "julio":"07","agosto":"08","septiembre":"09","setiembre":"09","octubre":"10",
            "noviembre":"11","diciembre":"12"}
    pattern=r"\b(\d{1,2})\s+de\s+("+"|".join(months)+r")\s+de\s+(\d{4})\b"
    text=re.sub(pattern,lambda m:f"{int(m.group(1)):02d}/{months[m.group(2).lower()]}/{m.group(3)}",text,flags=re.I)
    text=re.sub(r"\b(\d{4})-(\d{2})-(\d{2})\b",lambda m:f"{m.group(3)}/{m.group(2)}/{m.group(1)}",text)
    is_trasu=("TRASU" in str(ficha.get("tipo_acto") or "").upper() or
              "TRASU" in str(ficha.get("numero_acto") or "").upper() or
              "TRASU" in text[:180].upper())
    if is_trasu:
        notification=str(ficha.get("fecha_notificacion_emision") or "").strip()
        due=str(ficha.get("fecha_vencimiento") or "").strip()
        term=str(ficha.get("plazo_cumplimiento") or "").strip()
        obligation=str(ficha.get("obligacion_principal") or "").strip()
        if obligation:
            obligation=obligation.rstrip(" .;,")
            obligation=obligation[:1].lower()+obligation[1:]
        invalid={"","no identificado","pendiente de verificación","pendiente de verificacion"}
        if notification.lower() not in invalid and due.lower() not in invalid and term.lower() not in invalid:
            # The legal header is deterministic. Gemini may draft only the
            # evidentiary analysis that begins with "Al respecto".
            match=re.search(r"\bAl\s+respecto\b",text,flags=re.I)
            tail=text[match.start():].strip() if match else ""
            opening=(f"La Resolución TRASU fue notificada el {notification}, otorgando a la empresa operadora "
                     f"el plazo de {term} para cumplir la obligación principal consistente en {obligation}, "
                     f"plazo que vencía el {due}.")
            text=opening+(" "+tail if tail else "")
        else:
            raise ValueError("No se puede redactar un párrafo TRASU sin notificación, plazo y vencimiento verificados")
        folded_obligation=unicodedata.normalize("NFD",obligation.lower())
        folded_obligation="".join(c for c in folded_obligation if unicodedata.category(c)!="Mn")
        if any(k in folded_obligation for k in ("reconect","reactiv","restablec")):
            # Final deterministic guard against cross-matter language. Gemini
            # sometimes copies discount-specific wording even after the matter
            # filter. Such clauses are impossible in a reconnection analysis.
            text=re.sub(
                r"[,;]\s*(?:al\s+haber\s+)?(?:quedando|existiendo)?\s*(?:periodos?|meses?)\b[^.;]*?(?:descuent\w*|benefici\w*|importes?|montos?|registro\s+y\s+activaci[oó]n)[^.;]*[;.]?",
                ";",text,flags=re.I)
            text=re.sub(
                r",\s*toda\s+vez\s+que,?\s*(?:al\s+)?(?:haber\s+)?(?:quedar|quedando|existir|existiendo)\s+(?:periodos?|meses?)\b[^.]*?registro\s+y\s+activaci[oó]n[^.]*\.",
                ", toda vez que no existe una consulta o registro histórico con fecha objetiva que acredite que el servicio ya estaba activo o fue reconectado dentro del plazo otorgado.",
                text,flags=re.I)
            text=re.sub(
                r"[,;]\s*(?:al\s+)?(?:haber\s+)?(?:quedar|quedando|existir|existiendo)\s+(?:periodos?|meses?)\b[^.]*?registro\s+y\s+activaci[oó]n[^.]*[.;]?",
                ";",text,flags=re.I)
            text=re.sub(r";\s*;",";",text)
            text=re.sub(r"\s+([,;:.])",r"\1",text)
            # Sentence-level fallback: reject the whole contaminated sentence,
            # regardless of whether Gemini writes "quedando", "al quedar" or
            # another grammatical variant.
            clean_sentences=[]
            for sentence in re.split(r"(?<=[.!?])\s+",text):
                folded_sentence=unicodedata.normalize("NFD",sentence.lower())
                folded_sentence="".join(c for c in folded_sentence if unicodedata.category(c)!="Mn")
                wrong_period_language=(
                    any(k in folded_sentence for k in ("periodo","meses")) and
                    any(k in folded_sentence for k in ("registro y activacion","descuento","beneficio recurrente","nota de credito","importe","monto")))
                if not wrong_period_language: clean_sentences.append(sentence)
            text=" ".join(clean_sentences)
    # A completed paragraph must never jump from an empty "En consecuencia"
    # directly to subsanation. Restore the missing conclusion from the locked
    # evaluation result without asking the model to redraft anything.
    result_key=unicodedata.normalize("NFD",str(resultado or "").lower())
    result_key="".join(c for c in result_key if unicodedata.category(c)!="Mn")
    if result_key=="incumplio":
        conclusion="En consecuencia, al no haberse acreditado la ejecución oportuna del mandato, se habría configurado la infracción. Asimismo"
    elif result_key=="cumplio":
        conclusion="En consecuencia, se acreditó el cumplimiento oportuno del mandato. Asimismo"
    elif result_key=="inejecutable":
        conclusion="En consecuencia, se determinó que el mandato era inejecutable. Asimismo"
    else:
        conclusion="En consecuencia, no fue posible establecer una conclusión definitiva. Asimismo"
    text=re.sub(r"\bEn\s+consecuencia\s*[;,.]?\s*Asimismo\b",conclusion,text,flags=re.I)
    return text

def enforce_conditional_reconnection_timeline(result: dict[str,Any], extraction: dict[str,Any],
                                              paragraph: str, sources: dict[str,str]) -> str:
    """Reject the inference that a later active state proves the earlier condition or timely execution."""
    ficha=result.get("ficha") if isinstance(result.get("ficha"),dict) else {}
    obligation=_fold_legal_text(ficha.get("obligacion_principal"))
    selected_matter=str(sources.get("materia_criterios_seleccionada") or "")
    conditional_reconnection=(
        selected_matter=="Corte o Baja Injustificada" and
        any(k in obligation for k in ("reconect","reactiv","restablec")) and
        any(k in obligation for k in ("si ","suspend")))
    if not conditional_reconnection:
        return paragraph
    try:
        notification=parse_excel_date(ficha.get("fecha_notificacion_emision")).normalize()
        due=parse_excel_date(ficha.get("fecha_vencimiento")).normalize()
    except Exception:
        return paragraph

    timely_objective_proof=False
    for item in (extraction.get("medios_probatorios") or []):
        if not isinstance(item,dict): continue
        state=_fold_legal_text(item.get("estado"))
        date_text=str(item.get("fecha_ejecucion_acreditada") or "").strip()
        source_text=str(item.get("fuente_fecha_ejecucion") or item.get("cita") or "").strip()
        if state not in {"ejecutado","condicion_no_configurada"} or not date_text or not source_text:
            continue
        event_date=pd.to_datetime(date_text,dayfirst=True,errors="coerce")
        if pd.isna(event_date): continue
        event_date=pd.Timestamp(event_date).normalize()
        if state=="condicion_no_configurada" and event_date==notification:
            timely_objective_proof=True; break
        if state=="ejecutado" and notification<=event_date<=due:
            timely_objective_proof=True; break
    if timely_objective_proof:
        return paragraph

    result["resultado"]="Incumplió"
    result["clasificacion"]="PAS"
    result["subsanacion_voluntaria"]="no aplica"
    checklist=result.setdefault("evaluacion_juridica",{}).setdefault("checklist",[])
    checklist.append(
        "Validación temporal: no existe prueba objetiva fechada que acredite que la condición no se configuró "
        "en la fecha de notificación o que la reconexión se ejecutó hasta el vencimiento.")
    respect=re.search(r"\bAl\s+respecto\b.*?\.(?=\s+[A-ZÁÉÍÓÚÑ]|$)",str(paragraph or ""),flags=re.I|re.S)
    allegation=respect.group(0).strip() if respect else (
        "Al respecto, la empresa operadora señaló que el servicio se encontraba activo y operativo.")
    notification_text=notification.strftime("%d/%m/%Y")
    due_text=due.strftime("%d/%m/%Y")
    return (f"{allegation} Ahora bien, dicho medio acredita únicamente el estado del servicio en la fecha de "
            f"su consulta o comunicación, pero no contiene un registro histórico con fecha objetiva que demuestre que "
            f"el servicio ya estaba activo el {notification_text} ni que hubiese sido reconectado hasta el {due_text}. "
            "Por tanto, un estado posterior no permite inferir que la condición del mandato no se configuró al momento "
            "de la notificación. En consecuencia, al no haberse acreditado la ejecución oportuna del mandato, se habría "
            "configurado la infracción. Asimismo, no resulta aplicable el eximente de subsanación voluntaria, debido a "
            "que no se acreditó el cese total de la conducta ni la reversión integral de sus efectos.")

def ai_evaluate(payload: dict[str,Any]) -> dict[str,Any]:
    if payload.get("tipo_acto")=="Resolución TRASU":
        verified_date=str(payload.get("fecha_verificada") or "").strip()
        verified_due=str(payload.get("fecha_vencimiento") or "").strip()
        invalid={"","no identificado","pendiente de verificación","pendiente de verificacion"}
        if verified_date.lower() in invalid:
            raise ValueError("No se puede evaluar cumplimiento TRASU sin fecha de notificación verificada")
        if verified_due.lower() in invalid:
            raise ValueError("No se puede evaluar cumplimiento TRASU sin fecha de vencimiento calculada")
    template="PLANTILLAS cumplimiento.docx" if payload["tipo_acto"]=="Resolución TRASU" else "PLANTILLAS DENUNCIAS ACTUALIZADAS.docx"
    case_context=json.dumps(payload,ensure_ascii=False)
    verified_principal=payload.get("obligacion_y_plazo_principales_verificados") or {}
    obligation_for_criteria=(str(verified_principal.get("obligacion_principal") or "").strip()
                             if isinstance(verified_principal,dict) else "")
    if not obligation_for_criteria:
        obligation_for_criteria=str(payload.get("parte_resolutiva_trasu") or "").strip()
    sources=legal_sources(case_context,template,obligation_for_criteria)
    extraction_schema={"acto":{"numero":"","fecha":"","mandato_textual":""},"obligacion_extraida_parte_resolutiva":{"texto":"","articulos_numerales":[],"plazo_principal_textual":""},"obligaciones_accesorias_excluidas":[{"texto":"","plazo":"","motivo_exclusion":""}],"obligaciones":[{"componente":"","periodo":"","plazo_expreso":"","prueba_exigible":""}],"medios_probatorios":[{"documento":"","fecha_documento":"","fecha_ejecucion_acreditada":"","fuente_fecha_ejecucion":"","hecho_acreditado":"","cita":"","estado":"ejecutado|condicion_no_configurada|programado|en_curso|no_acreditado"}],"matriz_cumplimiento":[{"componente":"","estado":"acreditado|parcial|no_acreditado","sustento":""}],"datos_no_identificados":[]}
    extraction_system="""Actúa como extractor jurídico OSIPTEL. Separa hechos de conclusiones.

REGLA DE ORIGEN DE LA OBLIGACIÓN: si el acto es una Resolución TRASU, lee COMPLETA y CONJUNTAMENTE todos los artículos y numerales del campo parte_resolutiva_trasu. Identifica la obligación material principal relacionada con la prestación o controversia del servicio y su propio plazo, cualquiera sea el verbo o la forma jurídica empleados. La obligación y su plazo pueden estar en numerales consecutivos y no tienen que aparecer en el mismo numeral que declara fundado el reclamo. No uses listas cerradas de verbos. Copia el mandato con fidelidad y sepáralo por componentes, periodos, montos y condiciones.

REGLA OBLIGATORIA — VARIOS NUMERALES EN LA PARTE RESOLUTIVA: vincula cada plazo con la obligación concreta a la que corresponde. Evalúa solamente la ejecución material que resuelve la controversia. Excluye de la evaluación los numerales posteriores que ordenan comunicar o informar al usuario sobre el cumplimiento, informar o acreditar ante el TRASU, remitir constancias o realizar otro reporte formal, junto con sus respectivos plazos. Regístralos en obligaciones_accesorias_excluidas para conservar trazabilidad, pero nunca los mezcles con la obligación o el plazo principales ni los uses como fundamento de incumplimiento.

Los antecedentes y considerandos solo pueden aclarar una referencia o el alcance de un mandato ya contenido en HA RESUELTO. Nunca pueden crear, sustituir o ampliar la obligación. Las alegaciones, cartas posteriores y pruebas de ejecución tampoco pueden definirla.

Para cada prueba distingue ejecución efectiva de solicitud, programación, caso abierto o gestión en curso. Una afirmación de la empresa no prueba por sí sola el hecho. Cita el documento que respalda cada dato.

REGLA TEMPORAL OBLIGATORIA: distingue siempre la fecha del documento o de su remisión de la fecha efectiva de ejecución que el documento acredita. Un documento posterior al vencimiento puede acreditar una ejecución anterior únicamente cuando contiene un registro histórico, fecha de operación, evento de sistema u otro dato objetivo que identifique esa ejecución anterior. Un printer que solo muestra el estado actual "activo", sin fecha histórica de reconexión o de estado, no acredita por sí solo que la obligación se ejecutó dentro del plazo. Si el medio carece de fecha efectiva, deja fecha_ejecucion_acreditada vacía; nunca copies allí automáticamente la fecha de la carta.

RECONEXIÓN CONDICIONADA: si el mandato ordena reconectar solamente cuando el servicio se encuentre suspendido, verifica con evidencia fechada si la condición existía en la fecha de notificación y durante el plazo. Para concluir condicion_no_configurada debe existir un histórico o consulta fechada que acredite que el servicio ya estaba activo en esa fecha relevante. Que aparezca activo en una consulta posterior no demuestra por sí solo que nunca estuvo suspendido ni que se reconectó oportunamente.

INVENTARIO DE EVIDENCIA: distingue entre un medio no presentado y un medio presentado cuyo contenido no acredita el hecho o cuya fecha no es verificable. Si un archivo, índice o página identifica expresamente un "Histórico de cortes y reconexiones", "Consulta del estado del servicio" o printer equivalente, regístralo como presentado y está prohibido afirmar que la empresa "no lo remitió". Si sus datos no muestran una fecha útil, indica con precisión que fue presentado pero no permite verificar temporalmente la reconexión.

REGLA OBLIGATORIA para obligaciones de informar, brindar, remitir, entregar o trasladar información al usuario (aplica en especial a devoluciones en efectivo y comunicaciones equivalentes): una carta o correo simple no acredita por sí solo que el usuario recibió la información. Exige acuse, confirmación de recepción o entrega, cargo de notificación con fecha o equivalente. Si solo existe envío sin recepción acreditada, marca el componente como no_acreditado.

EXCEPCIÓN OBLIGATORIA — AJUSTES, ANULACIONES O DESCUENTOS EN LA FACTURACIÓN: cuando la obligación consiste en ajustar, anular o descontar un importe en la facturación (no en devolver dinero en efectivo), la ejecución se acredita con la captura de pantalla del sistema o el histórico del estado de cuenta que muestre que el ajuste coincide con el importe ordenado por el TRASU, conforme a los criterios de la materia "Facturación y cobro". En estos casos NO se exige acreditar que el usuario recibió una notificación o carta sobre el ajuste; dicha comunicación, si existe, es evidencia adicional pero no condición de cumplimiento. No confundas la obligación de ajustar la facturación con una obligación de informar al usuario.

No evalúes todavía PAS ni subsanación. Devuelve JSON conforme al esquema."""
    extraction=parse_json_response(gemini_text(extraction_system,json.dumps({"esquema":extraction_schema,"caso":payload},ensure_ascii=False),json_mode=True)) or {}
    if not isinstance(extraction,dict): raise ValueError("Gemini no devolvió una extracción jurídica válida")
    schema={"ficha":{"expediente":"","empresa_operadora":"","usuario_abonado":"","servicio":"","tipo_acto":"","numero_acto":"","fecha_notificacion_emision":"","plazo_cumplimiento":"","fecha_vencimiento":"","obligacion_principal":"","medios_probatorios":[]},"trazabilidad":[],"evaluacion_juridica":{"checklist":[],"sustento_breve":"","tipo_incumplimiento":""},"resultado":"Cumplió|Incumplió|Inejecutable","subsanacion_voluntaria":"aplica|no aplica|no corresponde","clasificacion":"PAS|NO PAS","parrafo_final":"","datos_pendientes":[]}
    system="""Eres analista jurídico senior de OSIPTEL. Usa SOLO la evidencia y fuentes entregadas. No inventes fechas, pruebas ni conclusiones. Lo faltante es 'No identificado' o 'Pendiente de verificación'.

JERARQUÍA DOCUMENTAL OBLIGATORIA:
1. Sigue literalmente instrucciones_juridicas.txt para determinar el tipo de acto, el orden del análisis, la prueba exigible y los datos que no pueden inferirse.
2. El sistema ya consultó criterios_evaluacion_obligatorios.txt y seleccionó, desde la obligación principal, la materia y las filas aplicables. Aplica íntegramente SOLO esas filas como instrucciones vinculantes. No uses filas no seleccionadas ni conocimiento general para sustituirlas. Conserva exactamente el significado de PAS, NO PAS, PLAZO e INEJECUTABLE de cada fila.
3. Aplica como reglas vinculantes las filas seleccionadas de PAUTAS PAS.xlsx, separando cumplimiento, razonabilidad, cese y subsanación voluntaria.
4. Redacta con la estructura y lenguaje de PLANTILLAS cumplimiento.docx cuando sea una Resolución TRASU.
5. Redacta con la estructura y lenguaje de PLANTILLAS DENUNCIAS ACTUALIZADAS.docx para cartas, denuncias, SAR, SARA o SAP.
6. Antes de concluir, identifica en el checklist el nombre del archivo y la fila o sección aplicada. Si no puedes identificar la regla utilizada, no emitas evaluación final y registra el dato como pendiente.
7. Está prohibido emitir una conclusión basada solamente en conocimiento general del modelo cuando contradiga cualquiera de esos documentos.

MÉTODO Y CONTROLES JURÍDICOS OBLIGATORIOS (en este orden):
0. Determina la materia y cita en el checklist las filas concretas de criterios y pautas usadas. No uses una regla de otra materia.
1. Evalúa cada componente y cada periodo del mandato. El cumplimiento parcial NO equivale a cumplimiento íntegro.
2. Determina la prueba exigible y el resultado únicamente con las filas seleccionadas de la materia aplicable y con los hechos acreditados.
3. Si el mandato contiene varios componentes, evalúa cada componente sin importar criterios de otra materia.
4. La subsanación voluntaria exige conjuntamente cese TOTAL de la conducta y reversión INTEGRAL de todos sus efectos antes del procedimiento. Que la materia no restrinja el servicio, por sí solo, jamás basta.
5. Si siguen componentes pendientes, explica la ausencia de cese total o reversión integral según las pautas seleccionadas.
6. Usa exclusivamente la fecha de vencimiento calculada y no la recalcules.
7. Aplica las instrucciones, criterios y pautas PAS entregados, incluso si contradicen una inferencia general del modelo. La plantilla determina la estructura de redacción.
8. Contrasta obligación por obligación con la matriz de pruebas; no generalices el resultado de un componente a los demás.
9. Mantén separadas tres decisiones: (a) cumplimiento material dentro del plazo, (b) razonabilidad para una ejecución tardía y (c) eximente de subsanación voluntaria. No conviertas automáticamente una ejecución tardía en subsanación.
10. No traslades pruebas, frases ni conclusiones de una materia distinta de la materia seleccionada.
11. La etiqueta NO PAS de una fila solo procede si todos los hechos descritos en esa fila están acreditados. Si faltan meses, extremos o prueba objetiva, aplica la fila específica de incumplimiento/PAS.
12. Antes de redactar, construye el checklist con: mandato; plazo; prueba exigible; prueba existente; periodos acreditados; periodos pendientes; regla aplicada; conclusión. Si existe contradicción, prevalece la evidencia documental y la regla específica.
13. El párrafo final debe indicar obligatoriamente y de forma expresa: fecha de notificación, número y tipo de días del plazo verificado, y fecha de vencimiento. Copia esos tres datos literalmente de la ficha; está prohibido recalcularlos u omitir el plazo.
14. Cada conclusión debe derivarse de hechos mencionados inmediatamente antes. No concluyas cumplimiento, incumplimiento, cese, reversión o subsanación si la matriz no identifica la prueba y los periodos que sustentan esa conclusión.
15. El párrafo final debe seguir literalmente la redacción y estructura de frases de la plantilla aplicable (PLANTILLAS cumplimiento.docx o PLANTILLAS DENUNCIAS ACTUALIZADAS.docx). Está prohibido agregar datos que la plantilla no contempla en esa oración, como el número de la resolución, carta, SARA, SAR, SAP o resolución de primera instancia, salvo que la plantilla lo incluya expresamente en su texto.
16. Está prohibido afirmar que una obligación se ejecutó "dentro del plazo" si la extracción probatoria no contiene una fecha_ejecucion_acreditada igual o anterior a la fecha de vencimiento. Distingue fecha de carta, fecha de captura y fecha histórica de ejecución. Para una reconexión condicionada, un estado "activo" consultado después del vencimiento solo acredita el estado en esa fecha posterior; no acredita la inexistencia de suspensión en la fecha de notificación ni una reconexión oportuna, salvo que el propio histórico identifique esas fechas.
17. CONTROL DE MATERIA: aplica únicamente criterios correspondientes a la obligación principal identificada. Si la obligación es reconectar, reactivar o acreditar operatividad, está prohibido fundamentar el análisis con reglas o expresiones sobre descuentos recurrentes, meses o periodos pendientes, importes, notas de crédito, registro o activación de beneficios, ofertas o promociones. Esas expresiones solo proceden cuando el mandato principal versa realmente sobre esas materias. Antes de redactar, elimina del razonamiento cualquier criterio perteneciente a una materia distinta.
18. CONTROL DE EXISTENCIA DOCUMENTAL: no confundas ausencia de un documento con insuficiencia de su contenido. Si el expediente o la extracción probatoria identifica que se presentó un histórico, consulta, printer, acta o constancia, menciona que fue presentado. Solo concluye que no acredita el cumplimiento cuando falte en ese medio la fecha, el evento o el dato objetivo exigible. Nunca escribas "no remitió" o "no adjuntó" respecto de un medio que aparece en el inventario documental.

Devuelve JSON válido conforme al esquema y un párrafo final completo, cronológico, con obligación, pruebas por periodo, contraste, conclusión y análisis separado de subsanación."""
    user=json.dumps({"esquema":schema,"caso":{k:v for k,v in payload.items() if k!="documentos"},"extraccion_probatoria":extraction,"fuentes":sources},ensure_ascii=False)
    result=parse_json_response(gemini_text(system,user,json_mode=True)) or {}
    if not isinstance(result,dict): raise ValueError("Gemini no devolvió una evaluación jurídica válida")
    if not isinstance(result.get("ficha"),dict): result["ficha"]={}
    if not isinstance(result.get("evaluacion_juridica"),dict): result["evaluacion_juridica"]={}
    checklist=result["evaluacion_juridica"].get("checklist")
    if not isinstance(checklist,list): checklist=[]
    selection_note=("criterios_evaluacion_obligatorios.txt — materia: "
                    f"{sources['materia_criterios_seleccionada']}; filas consultadas: "
                    f"{sources['filas_criterios_seleccionadas']}.")
    if not any("criterios_evaluacion_obligatorios.txt" in str(item) for item in checklist):
        checklist.insert(0,selection_note)
    result["evaluacion_juridica"]["checklist"]=checklist
    trace=result.get("trazabilidad")
    if not isinstance(trace,list): trace=[]
    trace.append(selection_note)
    result["trazabilidad"]=trace
    # Verified spreadsheet values are authoritative; the model cannot recalculate them.
    result.setdefault("ficha",{})["expediente"]=payload.get("expediente_detectado","No identificado")
    result["ficha"]["fecha_notificacion_emision"]=payload.get("fecha_verificada","No identificado")
    result["ficha"]["plazo_cumplimiento"]=payload.get("plazo_verificado","Pendiente de verificación")
    result["ficha"]["fecha_vencimiento"]=payload.get("fecha_vencimiento","Pendiente de verificación")
    verified_principal=payload.get("obligacion_y_plazo_principales_verificados") or {}
    resolutive_obligation=extraction.get("obligacion_extraida_parte_resolutiva") or {}
    if isinstance(verified_principal,dict) and str(verified_principal.get("obligacion_principal","")).strip():
        result["ficha"]["obligacion_principal"]=str(verified_principal["obligacion_principal"]).strip()
    elif isinstance(resolutive_obligation,dict) and str(resolutive_obligation.get("texto","")).strip():
        result["ficha"]["obligacion_principal"]=str(resolutive_obligation["texto"]).strip()
    elif payload.get("tipo_acto")=="Resolución TRASU":
        raise ValueError("Se leyó HA RESUELTO, pero no fue posible distinguir con seguridad la obligación principal de los deberes accesorios")
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
    if payload.get("tipo_acto")=="Resolución TRASU":
        # The applicable template starts with "La Resolución TRASU" and does
        # not reproduce the identifying number in the final paragraph.
        notification=str(payload.get("fecha_verificada")).strip()
        paragraph=re.sub(r"^La\s+Resoluci[oó]n(?:\s+TRASU)?(?:\s+N(?:ro\.?|[.°º])?\s*[0-9][A-Z0-9./-]*)?\s+fue\s+(?:emitida|notificada)(?:\s+en\s+fecha\s+no\s+identificada|\s+el\s+[^,.;]+)",
                         f"La Resolución TRASU fue notificada el {notification}",paragraph,count=1,flags=re.I)
        paragraph=re.sub(r"^La\s+Resoluci[oó]n(?:\s+TRASU)?\s+N(?:ro\.?|[.°º])?\s*[0-9][A-Z0-9./-]*",
                         "La Resolución TRASU",paragraph,count=1,flags=re.I)
    # The templates never cite the act's identifying number; strip it if the model added it anyway.
    numero_acto=str(result.get("ficha",{}).get("numero_acto","")).strip()
    if numero_acto and numero_acto.lower() not in {"no identificado","pendiente de verificación","pendiente de verificacion"}:
        paragraph=re.sub(r"\s*N\.?[°º]\s*"+re.escape(numero_acto),"",paragraph)
    # Avoid an accidental exact duplication of the generated paragraph.
    half=len(paragraph)//2
    if len(paragraph)%2==0 and paragraph[:half]==paragraph[half:]:
        paragraph=paragraph[:half].strip()
    paragraph=enforce_conditional_reconnection_timeline(result,extraction,paragraph,sources)
    result["parrafo_final"]=normalize_legal_paragraph(paragraph,result.get("ficha",{}),result.get("resultado",""))
    return result

def regenerate_paragraph(result: dict[str,Any]) -> str:
    context={"ficha":result.get("ficha",{}),"evaluacion_juridica":result.get("evaluacion_juridica",{}),"resultado":result.get("resultado"),"subsanacion_voluntaria":result.get("subsanacion_voluntaria"),"clasificacion":result.get("clasificacion"),"datos_pendientes":result.get("datos_pendientes",[]),"instrucciones":INSTRUCCIONES.read_text("utf-8",errors="ignore")[:40000]}
    response=gemini_text("Regenera únicamente el párrafo jurídico final con los datos aportados. Usa todas las fechas exclusivamente en formato dd/mm/aaaa. Si es TRASU, inicia con 'La Resolución TRASU fue notificada el dd/mm/aaaa' y no incluyas el número de resolución ni sustituyas la notificación por la emisión. No inventes ni completes faltantes. Devuelve JSON {parrafo_final:string}.",json.dumps(context,ensure_ascii=False),json_mode=True)
    paragraph=parse_json_response(response)["parrafo_final"]
    return normalize_legal_paragraph(paragraph,result.get("ficha",{}),result.get("resultado",""))

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
                # A document whose OCR/reading failed must never be treated as usable
                # content — otherwise the AI can latch onto an unrelated document and
                # fabricate an obligation instead of honestly reporting "not found".
                failed_docs=[n for n,t in texts.items() if str(t).startswith("[OCR no disponible") or str(t).startswith("[Error al leer")]
                failed_details=" | ".join(str(texts[n])[:1200] for n in failed_docs)
                for n in failed_docs: texts[n]=""
                if not texts or not any(str(x).strip() for x in texts.values()):
                    raise ValueError("No se pudo extraer texto de los archivos del expediente"+(f" (falló la lectura de: {', '.join(failed_docs)})" if failed_docs else ""))
                st.session_state.analysis_status="Documentos leídos; verificando expediente y plazo"
                combined="\n\n".join(f"### {n}\n{t}" for n,t in texts.items())
                document_types=[classify(n,t) for n,t in texts.items()]
                priority=["Resolución TRASU","SARA","SAR","SAP","Denuncia","Carta","Resolución de primera instancia"]
                unmistakable_trasu=(any("TRASU" in str(n).upper() for n in texts) or
                                    bool(re.search(r"\bTRASU\b",combined,re.I)))
                tipo="Resolución TRASU" if unmistakable_trasu else next((candidate for candidate in priority if candidate in document_types),"No identificado")
                resolutive=extract_resolutive_part(texts) if tipo=="Resolución TRASU" else None
                st.session_state["debug_resolutivo"]=resolutive
                if tipo=="Resolución TRASU" and not resolutive:
                    raise ValueError("No se identificó la sección HA RESUELTO completa; no es posible establecer la obligación principal sin esa sección"+(f". Detalle de lectura: {failed_details}" if failed_details else ""))
                principal_data=extract_trasu_mandate_and_term(resolutive) if resolutive else {}
                searchable="\n".join(u.name for u in uploads)+"\n"+combined
                expediente=parse_trasu_name(searchable) or identify_exact_expediente(searchable) or "No identificado"
                notice=None; due=None; term=None
                if tipo=="Resolución TRASU":
                    notice=exact_notification(expediente) if expediente!="No identificado" else None
                    if not notice:
                        st.session_state.analysis_error="No se encontró coincidencia exacta del expediente en notificaciones mayo.xlsx. Expediente detectado: "+expediente
                    else:
                        deadline=calculate_due(notice,str(principal_data.get("plazo_principal_textual") or ""))
                        if deadline: due,term=deadline
                        if not deadline: st.session_state.analysis_error="No se pudo determinar el plazo o calcular el vencimiento con CONTADOR DE PLAZOS - TRASU 2026.xlsx"
                if tipo!="Resolución TRASU" or (notice and due):
                    st.session_state.analysis_status="Aplicando instrucciones, criterios, pautas y plantillas"
                    payload={"tipo_acto":tipo,"expediente_detectado":expediente,"fecha_verificada":notice or "No identificado","plazo_verificado":term or "Pendiente de verificación","fecha_vencimiento":due or "Pendiente de verificación","parte_resolutiva_trasu":resolutive or "No corresponde","obligacion_y_plazo_principales_verificados":principal_data,"documentos":texts}
                    st.session_state.result=ai_evaluate(payload); st.session_state.texts=texts
                    st.session_state.analysis_status="Evaluación completada"
            except Exception as e:
                st.session_state.analysis_error=f"No se pudo completar el análisis: {type(e).__name__}: {e}"
                st.session_state.analysis_status=""
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

# CumpleTRASU

Dashboard LegalTech en Streamlit para evaluar el cumplimiento de actos de OSIPTEL/TRASU con fuentes permanentes y expedientes particulares. La aplicación no simula conclusiones: bloquea la evaluación si faltan las fuentes, una notificación TRASU exacta o un vencimiento calculable.

## Preparación

1. Instale Python 3.11 o 3.12.
2. Copie en `fuentes_permanentes/`, con sus nombres exactos:
   - `PLANTILLAS cumplimiento.docx`
   - `PLANTILLAS DENUNCIAS ACTUALIZADAS.docx`
   - `notificaciones mayo.xlsx`
   - `CONTADOR DE PLAZOS - TRASU 2026.xlsx`
   - `CRITERIOS DE EVALUACION DE CUMPLIMIENTO.xlsx`
   - `PAUTAS PAS.xlsx`
3. Reemplace el marcador de `instrucciones/instrucciones_juridicas.txt` por las instrucciones jurídicas reales de “PÁRRAFOS DE EVALUACIÓN”.
4. Duplique `.env.example` como `.env` y añada su clave de OpenAI. Nunca suba `.env` al repositorio.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Para OCR instale también Tesseract OCR con el paquete de idioma español. El OCR de PDF requiere Poppler instalado en el sistema. RAR puede requerir UnRAR. Si esos programas no están disponibles, la app seguirá leyendo PDF digital, Word, Excel, CSV y texto, y mostrará los errores de extracción sin inventar contenido.

## Uso

La persona usuaria carga solamente el expediente particular. Al pulsar “Analizar expediente con IA”, los archivos se guardan en una carpeta temporal, se extraen los formatos compatibles, se consultan las fuentes permanentes y se solicita una respuesta JSON estructurada. Los temporales se eliminan después del análisis. Las evaluaciones guardadas quedan en `salidas/evaluaciones.xlsx`.

En casos TRASU, la búsqueda en `NRO_EXPEDIENTE` es exacta y solo se toma `FEC_NOT_EMP_ELE_TEXTO`. Si no hay coincidencia, no se emite evaluación final.

## Streamlit Community Cloud

1. Cree un repositorio privado en GitHub y suba esta carpeta. Antes, confirme que las fuentes pueden almacenarse allí conforme a sus reglas de confidencialidad y protección de datos.
2. Entre a [Streamlit Community Cloud](https://share.streamlit.io/), seleccione **Create app**, el repositorio, la rama y `app.py`.
3. En **Advanced settings → Secrets**, configure:

```toml
OPENAI_API_KEY = "su-clave"
OPENAI_MODEL = "gpt-4.1-mini"
```

4. Despliegue la app y comparta el enlace generado. Para expedientes confidenciales se recomienda infraestructura privada, control de acceso y una política de retención; un enlace público no es apropiado.

Community Cloud puede no incluir Tesseract, Poppler o UnRAR. Para OCR y archivos comprimidos completos use un contenedor o servidor que permita instalar dependencias del sistema.

## Seguridad y límites

- No escriba claves en el código ni en archivos versionados.
- Verifique la autorización para enviar expedientes al proveedor de IA.
- Revise siempre la conclusión con un profesional jurídico.
- Las fuentes se leen desde disco en cada evaluación, por lo que pueden actualizarse sin cambiar el código.

import streamlit as st
import pandas as pd
import io
import json
from datetime import datetime
import os
from openai import OpenAI
import PyPDF2
import docx2txt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Configuración de la página
st.set_page_config(
    page_title="Sistema de Análisis de Admisión - SDT",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS profesionales y modernos CON CONTRASTE CORREGIDO
st.markdown("""
<style>
    /* Paleta de colores profesional para educación */
    :root {
        --primary-blue: #1E3A8A;
        --secondary-blue: #3B82F6;
        --accent-purple: #7C3AED;
        --success-green: #059669;
        --warning-orange: #F59E0B;
        --error-red: #DC2626;
        --bg-light: #F8FAFC;
        --text-dark: #1E293B;
        --text-medium: #64748B;
    }
    
    /* Fondo principal */
    .main {
        background: linear-gradient(180deg, #EFF6FF 0%, #FFFFFF 100%);
        padding: 2rem 1rem;
    }
    
    .stApp {
        background: #F1F5F9;
    }
    
    /* Header personalizado */
    .app-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(30, 58, 138, 0.2);
        text-align: center;
    }
    
    .app-title {
        font-size: 2.5rem;
        font-weight: 800;
        color: white;
        margin: 0;
        letter-spacing: -0.5px;
    }
    
    .app-subtitle {
        font-size: 1.1rem;
        color: #E0E7FF;
        margin-top: 0.5rem;
        font-weight: 400;
    }
    
    /* File uploader mejorado */
    div[data-testid="stFileUploader"] {
        background: white;
        border-radius: 12px;
        padding: 2rem;
        border: 2px dashed #3B82F6;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    
    div[data-testid="stFileUploader"]:hover {
        border-color: #1E3A8A;
        box-shadow: 0 8px 12px rgba(30, 58, 138, 0.1);
    }
    
    /* CORRECCIÓN: Labels del file uploader con mejor contraste */
    div[data-testid="stFileUploader"] label,
    div[data-testid="stFileUploader"] small {
        color: #1E293B !important;
        font-weight: 500 !important;
    }
    
    /* Botones mejorados */
    .stButton > button {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.3);
        transition: all 0.3s ease;
        letter-spacing: 0.3px;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(30, 58, 138, 0.4);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Métricas mejoradas CON CONTRASTE CORREGIDO */
    div[data-testid="metric-container"] {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        border-left: 4px solid #3B82F6;
    }
    
    div[data-testid="metric-container"] label {
        color: #475569 !important;
        font-size: 0.875rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #1E3A8A !important;
        font-size: 2rem;
        font-weight: 700;
    }
    
    /* Cards de información CON CONTRASTE MEJORADO */
    .info-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        margin: 1rem 0;
        border-left: 4px solid #3B82F6;
        transition: all 0.3s ease;
    }
    
    .info-card:hover {
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
        transform: translateY(-2px);
    }
    
    /* CORRECCIÓN: Texto en info-card con contraste adecuado */
    .info-card strong {
        color: #1E3A8A !important;
        font-size: 0.95rem;
        display: block;
        margin-bottom: 0.5rem;
    }
    
    .info-card-content {
        color: #1E293B !important;
        font-size: 1.1rem;
        font-weight: 600;
    }
    
    /* Score circle mejorado */
    .score-circle {
        display: inline-flex;
        width: 100px;
        height: 100px;
        border-radius: 50%;
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        align-items: center;
        justify-content: center;
        font-size: 2.2rem;
        font-weight: 800;
        box-shadow: 0 8px 20px rgba(30, 58, 138, 0.3);
        position: relative;
    }
    
    .score-circle::after {
        content: '/20';
        position: absolute;
        bottom: 10px;
        right: 10px;
        font-size: 0.7rem;
        opacity: 0.9;
        color: white;
    }
    
    /* Progress bar personalizada */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
    }
    
    /* Expander mejorado CON CONTRASTE */
    .streamlit-expanderHeader {
        background: white !important;
        border-radius: 8px;
        font-weight: 600;
        color: #1E3A8A !important;
        border: 1px solid #E2E8F0;
    }
    
    .streamlit-expanderHeader:hover {
        background: #F8FAFC !important;
        border-color: #3B82F6;
    }
    
    /* CORRECCIÓN: Contenido del expander con contraste */
    .streamlit-expanderContent {
        background: white;
        color: #1E293B !important;
    }
    
    /* Download button especial */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: white !important;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.3);
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(5, 150, 105, 0.4);
    }
    
    /* Alerts personalizados CON CONTRASTE */
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid;
    }
    
    /* CORRECCIÓN: Asegurar texto visible en alerts */
    .stSuccess, .stInfo, .stWarning, .stError {
        color: #1E293B !important;
    }
    
    /* Info boxes CON CONTRASTE MEJORADO */
    div[data-baseweb="notification"] {
        border-radius: 8px;
    }
    
    /* Badges de estado */
    .status-badge {
        display: inline-block;
        padding: 0.4rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.875rem;
        letter-spacing: 0.3px;
    }
    
    .badge-success {
        background: #D1FAE5;
        color: #065F46;
    }
    
    .badge-warning {
        background: #FEF3C7;
        color: #92400E;
    }
    
    .badge-error {
        background: #FEE2E2;
        color: #991B1B;
    }
    
    .badge-info {
        background: #DBEAFE;
        color: #1E40AF;
    }
    
    /* Títulos mejorados */
    h1 {
        color: #1E3A8A !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px !important;
    }
    
    h2 {
        color: #1E3A8A !important;
        font-weight: 700 !important;
        margin-top: 2rem !important;
    }
    
    h3 {
        color: #3B82F6 !important;
        font-weight: 600 !important;
    }
    
    h4 {
        color: #1E293B !important;
        font-weight: 600 !important;
    }
    
    /* CORRECCIÓN: Párrafos con contraste adecuado */
    p, li, span, div {
        color: #1E293B;
    }
    
    /* Separator line */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, #3B82F6, transparent);
        margin: 2rem 0;
    }
    
    /* Caption mejorado CON CONTRASTE */
    .stCaption {
        color: #475569 !important;
        font-size: 0.875rem !important;
    }
    
    /* Spinner personalizado */
    .stSpinner > div {
        border-top-color: #3B82F6 !important;
    }
    
    /* Tablas CON CONTRASTE */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
    }
    
    .dataframe th {
        background-color: #1E3A8A !important;
        color: white !important;
    }
    
    .dataframe td {
        color: #1E293B !important;
    }
    
    /* Container principal */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Success/Warning/Error boxes específicos CON CONTRASTE */
    .success-box {
        background: #D1FAE5;
        border-left: 4px solid #059669;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        color: #065F46 !important;
    }
    
    .warning-box {
        background: #FEF3C7;
        border-left: 4px solid #F59E0B;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        color: #92400E !important;
    }
    
    .info-box {
        background: #DBEAFE;
        border-left: 4px solid #3B82F6;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        color: #1E40AF !important;
    }
    
    /* CORRECCIÓN: Info boxes adicionales con contenido visible */
    .info-box strong {
        color: #1E3A8A !important;
    }
    
    /* CORRECCIÓN: Asegurar contraste en drag and drop */
    div[data-testid="stFileUploader"] > div {
        color: #1E293B !important;
    }
    
    /* CORRECCIÓN: Texto en spinner */
    .stSpinner > div + div {
        color: #475569 !important;
    }
    
    /* CORRECCIÓN: Asegurar markdown es visible */
    .stMarkdown {
        color: #1E293B !important;
    }
    
    /* CORRECCIÓN: Nombre del archivo subido en negro */
    div[data-testid="stFileUploader"] section {
        color: #1E293B !important;
    }
    
    div[data-testid="stFileUploader"] section div {
        color: #1E293B !important;
    }
    
    div[data-testid="stFileUploader"] section small {
        color: #475569 !important;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .app-title {
            font-size: 1.8rem;
        }
        
        .score-circle {
            width: 80px;
            height: 80px;
            font-size: 1.8rem;
        }
        
        .block-container {
            padding-top: 1rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# Inicializar cliente OpenAI
@st.cache_resource
def get_openai_client():
    api_key = None
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except:
        api_key = os.getenv('OPENAI_API_KEY')
    
    if not api_key:
        st.error("⚠️ No se encontró la clave API de OpenAI")
        st.info("**Para uso local:** Configura OPENAI_API_KEY en .streamlit/secrets.toml")
        st.info("**Para Streamlit Cloud:** Configura OPENAI_API_KEY en Settings → Secrets")
        st.stop()
    return OpenAI(api_key=api_key)

client = get_openai_client()

# Funciones de extracción de texto
def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error al leer PDF: {str(e)}")
        return None

def extract_text_from_docx(file):
    try:
        text = docx2txt.process(file)
        return text
    except Exception as e:
        st.error(f"Error al leer DOCX: {str(e)}")
        return None

def extract_text_from_txt(file):
    try:
        return file.read().decode('utf-8')
    except Exception as e:
        st.error(f"Error al leer TXT: {str(e)}")
        return None

def read_excel_file(file):
    try:
        file_extension = file.name.split('.')[-1].lower()
        if file_extension == 'csv':
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error al leer archivo Excel/CSV: {str(e)}")
        return None

# Función principal de análisis con OpenAI
def analyze_admission_form(text_content):
    system_prompt = """ROL Y CONTEXTO:
Actúa como experto en Psicología Educativa y Motivación, especializado en la Teoría de la Autodeterminación (Self-Determination Theory, SDT) de Ryan y Deci.

MARCO TEÓRICO - TEORÍA DE LA AUTODETERMINACIÓN (SDT):
Evalúa la motivación académica basándote en:
- Motivación intrínseca: interés genuino, disfrute, curiosidad
- Regulaciones extrínsecas:
  * Externa: recompensas/presiones externas
  * Introyectada: presión interna, culpa, orgullo
  * Identificada: utilidad para metas personales
  * Integrada: coherencia con identidad y valores
- Amotivación: sin razón clara o desinterés
- Necesidades psicológicas básicas: autonomía, competencia, relación

CONTEXTO INSTITUCIONAL:
Universidad Continental - Modalidad a Distancia
Población diversa: todo Perú, 18+ años, muchos trabajan y tienen familia
Objetivo: diagnosticar niveles motivacionales, NO filtrar

CRITERIOS DE EVALUACIÓN (Escala 1-6):
Cada respuesta se evalúa según nivel motivacional:
6 = Motivación Intrínseca (interés genuino, disfrute)
5 = Motivación Integrada (coherencia con identidad y valores)
4 = Motivación Identificada (utilidad personal significativa)
3 = Motivación Introyectada (presión interna, orgullo, culpa)
2 = Motivación Extrínseca (recompensas/presiones externas)
1 = Amotivación (sin razón clara, desinterés)

ESTRUCTURA DE ANÁLISIS:
Debes analizar el formulario buscando evidencia de estos niveles motivacionales y responder SOLO con un objeto JSON válido (sin markdown, sin ```json) con esta estructura:

{
  "informacion_extraida": {
    "nombre": "...",
    "edad": "...",
    "programa": "...",
    "otros_campos": {}
  },
  "evaluacion_motivacional": {
    "eleccion_carrera": {
      "puntaje": 1-6,
      "tipo_motivacion": "Intrínseca/Integrada/Identificada/Introyectada/Extrínseca/Amotivación",
      "justificacion": "Evidencia textual breve (1-2 líneas)"
    },
    "experiencia_relacionada": {
      "puntaje": 1-6,
      "tipo_motivacion": "...",
      "justificacion": "..."
    },
    "uso_futuro": {
      "puntaje": 1-6,
      "tipo_motivacion": "...",
      "justificacion": "..."
    }
  },
  "necesidades_psicologicas": {
    "autonomia": "Alta/Media/Baja - Breve análisis",
    "competencia": "Alta/Media/Baja - Breve análisis",
    "relacion": "Alta/Media/Baja - Breve análisis"
  },
  "calificacion_real": "suma de puntajes (máx 18)",
  "calificacion_sobre_20": "calificación_real/18 * 20 (dos decimales)",
  "recomendaciones": "Basadas en SDT: fortalecer autonomía, competencia o relación según necesidad",
  "nivel_motivacional_general": "Predominantemente Intrínseco/Integrado/Identificado/Introyectado/Extrínseco/Amotivado"
}

IMPORTANTE:
- Evalúa SOLO la motivación expresada, NO la ortografía
- Sé objetivo, claro y no valorativo
- Usa evidencia textual del formulario
- Responde ÚNICAMENTE con el JSON (sin markdown ni bloques de código)
- Si hay campos faltantes, usa "N/A" o infiere del contexto
- Calcula calificacion_sobre_20 = (calificacion_real / 18) * 20
- Redondea a 2 decimales
"""

    user_prompt = f"""Analiza el siguiente formulario de admisión y proporciona un análisis completo según SDT:

{text_content}

Recuerda: responde ÚNICAMENTE con el JSON, sin ningún texto adicional ni bloques de código markdown."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,
            max_tokens=2000
        )
        
        content = response.choices[0].message.content.strip()
        
        if content.startswith('```'):
            content = content.split('```')[1]
            if content.startswith('json'):
                content = content[4:]
            content = content.strip()
        
        result = json.loads(content)
        result['tokens_used'] = response.usage.total_tokens
        result['timestamp'] = datetime.now().isoformat()
        result['success'] = True
        
        return result
        
    except json.JSONDecodeError as e:
        st.error(f"Error al parsear respuesta JSON: {str(e)}")
        return {"success": False, "error": "Error al parsear respuesta de IA"}
    except Exception as e:
        st.error(f"Error en análisis con OpenAI: {str(e)}")
        return {"success": False, "error": str(e)}

# Función para procesar registros de Excel
def process_excel_records(df, progress_bar, status_text):
    results = []
    total = len(df)
    
    required_fields = ['Respuesta 1', 'Respuesta 2', 'Respuesta 3']
    
    for idx, row in df.iterrows():
        # Actualizar UI con mejor contraste
        status_text.markdown(f"<div style='color: #1E3A8A; font-weight: 600;'>🔄 Procesando: Registro {idx + 1} de {total}</div>", unsafe_allow_html=True)
        progress_bar.progress((idx + 1) / total)
        
        # Extraer datos del Excel (usar los nombres de columnas correctos)
        apellidos_excel = row.get('Apellido(s)', row.get('Apellidos', ''))
        nombre_excel = row.get('Nombre', '')
        correo_excel = row.get('Dirección de correo', row.get('Correo electrónico', ''))
        edad_excel = row.get('Edad', '')
        programa_excel = row.get('Programa', row.get('Carrera', ''))
        
        form_text = f"""
Apellido(s): {apellidos_excel}
Nombre: {nombre_excel}
Correo: {correo_excel}
Edad: {edad_excel}
Programa: {programa_excel}

Pregunta 1: ¿Qué características de esta carrera llamaron tu atención?
Respuesta 1: {row.get('Respuesta 1', 'Sin respuesta')}

Pregunta 2: Experiencia relacionada con la carrera
Respuesta 2: {row.get('Respuesta 2', 'Sin respuesta')}

Pregunta 3: ¿Cómo aplicarías lo aprendido?
Respuesta 3: {row.get('Respuesta 3', 'Sin respuesta')}
"""
        
        missing_fields = [field for field in required_fields if pd.isna(row.get(field)) or str(row.get(field)).strip() == '']
        
        if missing_fields:
            results.append({
                'success': False,
                'registro_numero': idx + 1,
                'nombre': nombre_excel,
                'apellidos': apellidos_excel,
                'correo': correo_excel,
                'error': f"Campos faltantes: {', '.join(missing_fields)}"
            })
            continue
        
        analysis = analyze_admission_form(form_text)
        
        # Usar datos del Excel directamente (como hace server.js)
        result = {
            'registro_numero': idx + 1,
            'apellidos': apellidos_excel,
            'nombre': nombre_excel,
            'correo': correo_excel,
            'success': analysis.get('success', False),
        }
        
        if analysis.get('success'):
            result['analysis'] = analysis
        else:
            result['error'] = analysis.get('error', 'Error desconocido')
        
        results.append(result)
    
    return results

# Función para generar Excel de resultados
def generate_excel_report(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados Análisis SDT"
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    headers = [
        'N°', 'Apellido(s)', 'Nombre', 'Dirección de correo', 'Calif. Real', 'Calif. /20',
        'R1 Punt.', 'R1 Justificación', 'R1 Tipo',
        'R2 Punt.', 'R2 Justificación', 'R2 Tipo',
        'R3 Punt.', 'R3 Justificación', 'R3 Tipo',
        'Nivel General', 'Autonomía', 'Competencia', 'Relación', 'Recomendaciones'
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    
    column_widths = [5, 15, 15, 25, 10, 10, 8, 40, 15, 8, 40, 15, 8, 40, 15, 20, 30, 30, 30, 50]
    for col_num, width in enumerate(column_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = width
    
    for result in results:
        r = result
        a = r.get('analysis', {})
        
        ws.append([
            r.get('registro_numero', ''),
            r.get('apellidos', ''),
            r.get('nombre', ''),
            r.get('correo', ''),
            a.get('calificacion_real', ''),
            a.get('calificacion_sobre_20', ''),
            a.get('evaluacion_motivacional', {}).get('eleccion_carrera', {}).get('puntaje', ''),
            a.get('evaluacion_motivacional', {}).get('eleccion_carrera', {}).get('justificacion', ''),
            a.get('evaluacion_motivacional', {}).get('eleccion_carrera', {}).get('tipo_motivacion', ''),
            a.get('evaluacion_motivacional', {}).get('experiencia_relacionada', {}).get('puntaje', ''),
            a.get('evaluacion_motivacional', {}).get('experiencia_relacionada', {}).get('justificacion', ''),
            a.get('evaluacion_motivacional', {}).get('experiencia_relacionada', {}).get('tipo_motivacion', ''),
            a.get('evaluacion_motivacional', {}).get('uso_futuro', {}).get('puntaje', ''),
            a.get('evaluacion_motivacional', {}).get('uso_futuro', {}).get('justificacion', ''),
            a.get('evaluacion_motivacional', {}).get('uso_futuro', {}).get('tipo_motivacion', ''),
            a.get('nivel_motivacional_general', ''),
            a.get('necesidades_psicologicas', {}).get('autonomia', ''),
            a.get('necesidades_psicologicas', {}).get('competencia', ''),
            a.get('necesidades_psicologicas', {}).get('relacion', ''),
            a.get('recomendaciones', '')
        ])
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 30
        for cell in row:
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = border
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# INTERFAZ PRINCIPAL
def main():
    # Header profesional
    st.markdown("""
    <div class="app-header">
        <h1 class="app-title">🎓 Sistema de Análisis de Admisión</h1>
        <p class="app-subtitle">Análisis Motivacional basado en la Teoría de la Autodeterminación (SDT)</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Información de uso CON MEJOR CONTRASTE
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;'>
            <p style='margin: 0; color: #475569; font-size: 0.95rem;'>
                <strong style='color: #1E3A8A;'>📄 Análisis Individual:</strong> PDF, DOCX, TXT | 
                <strong style='color: #1E3A8A;'>📊 Análisis Masivo:</strong> XLSX, XLS, CSV
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Upload de archivo
    uploaded_file = st.file_uploader(
        "📎 Seleccionar Archivo",
        type=['pdf', 'docx', 'doc', 'txt', 'xlsx', 'xls', 'csv'],
        help="Arrastra el archivo o haz clic para seleccionar",
        label_visibility="collapsed"
    )
    
    if uploaded_file:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        file_size = uploaded_file.size / (1024 * 1024)
        is_batch = file_extension in ['xlsx', 'xls', 'csv']
        
        # Info del archivo CON MEJOR CONTRASTE
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"""
            <div class='info-box'>
                <strong>Archivo:</strong> {uploaded_file.name}
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class='info-box'>
                <strong>Tamaño:</strong> {file_size:.2f} MB
            </div>
            """, unsafe_allow_html=True)
        with col3:
            mode_badge = "badge-info" if is_batch else "badge-success"
            mode_text = "Modo Masivo" if is_batch else "Modo Individual"
            st.markdown(f"""
            <div class='info-box'>
                <span class='status-badge {mode_badge}'>{mode_text}</span>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Botón de análisis centrado
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            analyze_button = st.button("🚀 Iniciar Análisis", type="primary", use_container_width=True)
        
        if analyze_button:
            with st.spinner("🔄 Procesando análisis..."):
                if is_batch:
                    # PROCESAMIENTO MASIVO
                    df = read_excel_file(uploaded_file)
                    
                    if df is not None:
                        st.success(f"✅ {len(df)} registros detectados")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        results = process_excel_records(df, progress_bar, status_text)
                        
                        status_text.markdown("<div style='color: #059669; font-weight: 600;'>✅ Análisis completado exitosamente</div>", unsafe_allow_html=True)
                        progress_bar.progress(1.0)
                        
                        st.session_state['batch_results'] = results
                        st.session_state['batch_filename'] = uploaded_file.name
                        
                        # Separador
                        st.markdown("<hr>", unsafe_allow_html=True)
                        
                        # Resultados
                        st.markdown("## 📊 Resultados del Análisis Masivo")
                        
                        success_count = sum(1 for r in results if r.get('success'))
                        avg_score = sum(
                            float(r['analysis']['calificacion_sobre_20']) 
                            for r in results 
                            if r.get('success') and r.get('analysis', {}).get('calificacion_sobre_20')
                        ) / success_count if success_count > 0 else 0
                        
                        # Métricas
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("📋 Total Registros", len(results))
                        with col2:
                            st.metric("✅ Exitosos", success_count)
                        with col3:
                            st.metric("⚠️ Con Errores", len(results) - success_count)
                        with col4:
                            st.metric("📈 Promedio", f"{avg_score:.2f}/20")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        # Botón de descarga
                        excel_buffer = generate_excel_report(results)
                        col1, col2, col3 = st.columns([1, 2, 1])
                        with col2:
                            st.download_button(
                                label="📥 Descargar Reporte Completo (Excel)",
                                data=excel_buffer,
                                file_name=f"Analisis_SDT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        # Detalle por postulante
                        st.markdown("### 👥 Detalle por Postulante")
                        
                        for i, result in enumerate(results):
                            with st.expander(
                                f"**{result['registro_numero']}. {result['apellidos']}, {result['nombre']}** • {result['correo']}",
                                expanded=False
                            ):
                                if result.get('success'):
                                    analysis = result['analysis']
                                    
                                    col1, col2 = st.columns([1, 3])
                                    with col1:
                                        st.markdown(f"""
                                        <div style='text-align: center; padding: 1rem;'>
                                            <div class='score-circle'>{analysis.get('calificacion_sobre_20', 'N/A')}</div>
                                        </div>
                                        """, unsafe_allow_html=True)
                                    with col2:
                                        st.markdown(f"<div style='color: #1E293B;'><strong style='color: #1E3A8A;'>📊 Calificación Real:</strong> {analysis.get('calificacion_real', 'N/A')}/18</div>", unsafe_allow_html=True)
                                        nivel = analysis.get('nivel_motivacional_general', 'N/A')
                                        st.markdown(f"<div style='color: #1E293B;'><strong style='color: #1E3A8A;'>🎯 Nivel Motivacional:</strong> {nivel}</div>", unsafe_allow_html=True)
                                    
                                    st.markdown("---")
                                    
                                    # Evaluación motivacional
                                    st.markdown("<h4 style='color: #1E3A8A;'>📝 Evaluación Motivacional Detallada</h4>", unsafe_allow_html=True)
                                    eval_mot = analysis.get('evaluacion_motivacional', {})
                                    
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        if 'eleccion_carrera' in eval_mot:
                                            st.info(f"**R1: Elección de Carrera**\n\n{eval_mot['eleccion_carrera'].get('puntaje')}/6 • {eval_mot['eleccion_carrera'].get('tipo_motivacion')}")
                                    with col2:
                                        if 'experiencia_relacionada' in eval_mot:
                                            st.info(f"**R2: Experiencia**\n\n{eval_mot['experiencia_relacionada'].get('puntaje')}/6 • {eval_mot['experiencia_relacionada'].get('tipo_motivacion')}")
                                    with col3:
                                        if 'uso_futuro' in eval_mot:
                                            st.info(f"**R3: Proyección**\n\n{eval_mot['uso_futuro'].get('puntaje')}/6 • {eval_mot['uso_futuro'].get('tipo_motivacion')}")
                                    
                                    # Necesidades psicológicas
                                    if 'necesidades_psicologicas' in analysis:
                                        st.markdown("<h4 style='color: #1E3A8A;'>🧠 Necesidades Psicológicas (SDT)</h4>", unsafe_allow_html=True)
                                        nec = analysis['necesidades_psicologicas']
                                        col1, col2, col3 = st.columns(3)
                                        with col1:
                                            st.success(f"**Autonomía**\n\n{nec.get('autonomia', 'N/A')}")
                                        with col2:
                                            st.success(f"**Competencia**\n\n{nec.get('competencia', 'N/A')}")
                                        with col3:
                                            st.success(f"**Relación**\n\n{nec.get('relacion', 'N/A')}")
                                    
                                    # Recomendaciones
                                    if 'recomendaciones' in analysis:
                                        st.markdown("<h4 style='color: #1E3A8A;'>💡 Recomendaciones Pedagógicas</h4>", unsafe_allow_html=True)
                                        st.info(analysis['recomendaciones'])
                                else:
                                    st.error(f"❌ **Error:** {result.get('error', 'Error desconocido')}")
                
                else:
                    # PROCESAMIENTO INDIVIDUAL
                    text_content = None
                    
                    if file_extension == 'pdf':
                        text_content = extract_text_from_pdf(uploaded_file)
                    elif file_extension in ['docx', 'doc']:
                        text_content = extract_text_from_docx(uploaded_file)
                    elif file_extension == 'txt':
                        text_content = extract_text_from_txt(uploaded_file)
                    
                    if text_content and text_content.strip():
                        st.success(f"✅ Texto extraído correctamente ({len(text_content)} caracteres)")
                        
                        analysis = analyze_admission_form(text_content)
                        
                        if analysis.get('success'):
                            # Separador
                            st.markdown("<hr>", unsafe_allow_html=True)
                            
                            st.markdown("## 📊 Resultado del Análisis Individual")
                            
                            # Información básica
                            info = analysis.get('informacion_extraida', {})
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.info(f"**👤 Nombre**\n\n{info.get('nombre', 'N/A')}")
                            with col2:
                                st.info(f"**📅 Edad**\n\n{info.get('edad', 'N/A')}")
                            with col3:
                                st.info(f"**🎓 Programa**\n\n{info.get('programa', 'N/A')}")
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            # Calificaciones
                            col1, col2 = st.columns([1, 2])
                            with col1:
                                st.markdown(f"""
                                <div style='text-align: center; padding: 2rem;'>
                                    <div class='score-circle'>{analysis.get('calificacion_sobre_20', 'N/A')}</div>
                                    <p style='color: #475569; margin-top: 1rem; font-weight: 600;'>Calificación Final</p>
                                </div>
                                """, unsafe_allow_html=True)
                            with col2:
                                st.metric("📊 Calificación Real", f"{analysis.get('calificacion_real', 'N/A')}/18", 
                                         help="Suma de los 3 puntajes (máximo 18)")
                                st.metric("🎯 Nivel Motivacional", analysis.get('nivel_motivacional_general', 'N/A'),
                                         help="Perfil predominante según SDT")
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            # Evaluación motivacional
                            st.markdown("### 📝 Evaluación Motivacional Detallada")
                            eval_mot = analysis.get('evaluacion_motivacional', {})
                            
                            for key, label, icon in [
                                ('eleccion_carrera', 'Elección de Carrera', '🎯'),
                                ('experiencia_relacionada', 'Experiencia Relacionada', '📚'),
                                ('uso_futuro', 'Proyección Futura', '🚀')
                            ]:
                                if key in eval_mot:
                                    item = eval_mot[key]
                                    with st.expander(f"{icon} **{label}** • Puntaje: {item.get('puntaje')}/6 • Tipo: {item.get('tipo_motivacion')}", expanded=True):
                                        st.markdown(f"<div style='color: #1E293B;'><strong style='color: #1E3A8A;'>Justificación:</strong> {item.get('justificacion')}</div>", unsafe_allow_html=True)
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            # Necesidades psicológicas
                            if 'necesidades_psicologicas' in analysis:
                                st.markdown("### 🧠 Análisis de Necesidades Psicológicas (SDT)")
                                nec = analysis['necesidades_psicologicas']
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.success(f"**🎯 Autonomía**\n\n{nec.get('autonomia', 'N/A')}")
                                with col2:
                                    st.success(f"**💪 Competencia**\n\n{nec.get('competencia', 'N/A')}")
                                with col3:
                                    st.success(f"**🤝 Relación**\n\n{nec.get('relacion', 'N/A')}")
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            # Recomendaciones
                            if 'recomendaciones' in analysis:
                                st.markdown("### 💡 Recomendaciones Pedagógicas")
                                st.info(analysis['recomendaciones'])
                            
                            # Metadata
                            st.markdown("<hr>", unsafe_allow_html=True)
                            st.caption(f"📄 **Archivo:** {uploaded_file.name} | 🔢 **Tokens:** {analysis.get('tokens_used', 'N/A')} | ⏰ **Procesado:** {datetime.fromisoformat(analysis.get('timestamp')).strftime('%d/%m/%Y %H:%M:%S')}")
                        
                        else:
                            st.error(f"❌ **Error en el análisis:** {analysis.get('error', 'Error desconocido')}")
                    else:
                        st.error("❌ No se pudo extraer texto del archivo o el archivo está vacío")

if __name__ == "__main__":
    main()

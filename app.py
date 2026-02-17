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

st.set_page_config(
    page_title="Sistema de Análisis de Admisión - SDT",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* ═══════════════════════════════════════════════════════════
       VARIABLES DE COLOR - CONTRASTE ÓPTIMO
       ═══════════════════════════════════════════════════════════ */
    :root {
        --primary-blue: #1E3A8A;
        --secondary-blue: #3B82F6;
        --accent-purple: #7C3AED;
        --success-green: #059669;
        --warning-orange: #F59E0B;
        --error-red: #DC2626;
        --bg-light: #F8FAFC;
        --text-dark: #0F172A;
        --text-medium: #334155;
        --text-light: #64748B;
        --border-color: #E2E8F0;
    }
    
    /* ═══════════════════════════════════════════════════════════
       FORZAR TEMA CLARO EN TODO EL APP
       ═══════════════════════════════════════════════════════════ */
    .main,
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    [data-testid="stVerticalBlock"],
    [data-testid="stAppViewBlockContainer"],
    section[data-testid="stSidebar"],
    .block-container {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
    }
    
    .main {
        background: linear-gradient(180deg, #EFF6FF 0%, #F1F5F9 100%) !important;
        padding: 2rem 1rem;
    }
    
    /* ═══════════════════════════════════════════════════════════
       FILE UPLOADER - FONDO CLARO (FIX #1)
       ═══════════════════════════════════════════════════════════ */
    div[data-testid="stFileUploader"] {
        background: #FFFFFF !important;
        border-radius: 12px;
        padding: 2rem;
        border: 2px dashed #3B82F6 !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    
    div[data-testid="stFileUploader"]:hover {
        border-color: #1E3A8A !important;
        box-shadow: 0 8px 12px rgba(30, 58, 138, 0.1);
    }
    
    /* Todos los textos dentro del uploader */
    div[data-testid="stFileUploader"] label,
    div[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] span,
    div[data-testid="stFileUploader"] p,
    div[data-testid="stFileUploader"] div {
        color: #0F172A !important;
        font-weight: 500 !important;
    }
    
    /* Zona de drop interna */
    div[data-testid="stFileUploader"] section,
    div[data-testid="stFileUploader"] section > div,
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"],
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: #0F172A !important;
        border-color: #3B82F6 !important;
    }
    
    /* Botón Browse files */
    div[data-testid="stFileUploader"] button,
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #3B82F6 !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: 600 !important;
    }
    
    div[data-testid="stFileUploader"] button:hover,
    [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #1E3A8A !important;
    }
    
    /* Drag and drop text */
    [data-testid="stFileUploaderDropzoneInstructions"] div,
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: #334155 !important;
    }
    
    /* Archivo subido - nombre y tamaño */
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] * {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    
    /* ═══════════════════════════════════════════════════════════
       HEADER
       ═══════════════════════════════════════════════════════════ */
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
        color: #FFFFFF !important;
        margin: 0;
        letter-spacing: -0.5px;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .app-subtitle {
        font-size: 1.1rem;
        color: #F1F5F9 !important;
        margin-top: 0.5rem;
        font-weight: 400;
    }
    
    /* ═══════════════════════════════════════════════════════════
       BOTONES
       ═══════════════════════════════════════════════════════════ */
    .stButton > button {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: #FFFFFF !important;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.3);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(30, 58, 138, 0.4);
    }
    
    /* ═══════════════════════════════════════════════════════════
       MÉTRICAS
       ═══════════════════════════════════════════════════════════ */
    div[data-testid="metric-container"] {
        background: #FFFFFF !important;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        border-left: 4px solid #3B82F6;
    }
    
    div[data-testid="metric-container"] label {
        color: #334155 !important;
        font-size: 0.875rem;
        font-weight: 600;
        text-transform: uppercase;
    }
    
    div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #1E3A8A !important;
        font-size: 2rem;
        font-weight: 700;
    }
    
    /* ═══════════════════════════════════════════════════════════
       INFO CARDS
       ═══════════════════════════════════════════════════════════ */
    .info-card {
        background: #FFFFFF;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        margin: 1rem 0;
        border-left: 4px solid #3B82F6;
    }
    
    .info-card strong { color: #1E3A8A !important; }
    .info-card p, .info-card span, .info-card div { color: #0F172A !important; }
    
    /* Score circle */
    .score-circle {
        display: inline-flex;
        width: 100px;
        height: 100px;
        border-radius: 50%;
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: #FFFFFF !important;
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
        color: #FFFFFF !important;
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
    }
    
    /* ═══════════════════════════════════════════════════════════
       EXPANDERS - FIX #3: FONDO CLARO AL DESPLEGAR
       ═══════════════════════════════════════════════════════════ */
    /* Contenedor general del expander */
    div[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        margin-bottom: 0.5rem;
        overflow: hidden;
    }
    
    /* Header del expander - cerrado */
    div[data-testid="stExpander"] details > summary {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-weight: 600;
        padding: 0.75rem 1rem;
        border: none !important;
    }
    
    div[data-testid="stExpander"] details > summary:hover {
        background-color: #F8FAFC !important;
    }
    
    /* Header del expander - abierto */
    div[data-testid="stExpander"] details[open] > summary {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border-bottom: 2px solid #3B82F6 !important;
    }
    
    /* Texto dentro del header */
    div[data-testid="stExpander"] details > summary *,
    div[data-testid="stExpander"] details > summary span,
    div[data-testid="stExpander"] details > summary p,
    div[data-testid="stExpander"] summary [data-testid="stMarkdownContainer"],
    div[data-testid="stExpander"] summary [data-testid="stMarkdownContainer"] * {
        color: #0F172A !important;
    }
    
    /* Contenido del expander - abierto (FIX PRINCIPAL) */
    div[data-testid="stExpander"] details > div,
    div[data-testid="stExpander"] details > div > div,
    div[data-testid="stExpander"] details[open] > div,
    div[data-testid="stExpander"] details[open] > div > div,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"],
    [data-testid="stExpanderDetails"] {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: #0F172A !important;
    }
    
    /* Todo texto dentro del expander abierto */
    div[data-testid="stExpander"] details[open] p,
    div[data-testid="stExpander"] details[open] div,
    div[data-testid="stExpander"] details[open] span,
    div[data-testid="stExpander"] details[open] strong,
    div[data-testid="stExpander"] details[open] li,
    [data-testid="stExpanderDetails"] *,
    [data-testid="stExpanderDetails"] p,
    [data-testid="stExpanderDetails"] div,
    [data-testid="stExpanderDetails"] span {
        color: #0F172A !important;
        background-color: transparent !important;
    }
    
    /* Alerts dentro de expanders */
    div[data-testid="stExpander"] [data-testid="stAlert"],
    [data-testid="stExpanderDetails"] [data-testid="stAlert"] {
        background-color: #DBEAFE !important;
    }
    
    div[data-testid="stExpander"] [data-testid="stAlert"] *,
    [data-testid="stExpanderDetails"] [data-testid="stAlert"] * {
        color: #1E3A8A !important;
    }
    
    /* ═══════════════════════════════════════════════════════════
       DOWNLOAD BUTTON
       ═══════════════════════════════════════════════════════════ */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: #FFFFFF !important;
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
    
    /* ═══════════════════════════════════════════════════════════
       ALERTS
       ═══════════════════════════════════════════════════════════ */
    .stAlert { border-radius: 8px; border-left: 4px solid; }
    
    /* Success alerts */
    [data-testid="stAlert"][data-baseweb*="positive"],
    div[data-baseweb="notification"][kind="positive"],
    .element-container .stSuccess {
        background-color: #D1FAE5 !important;
    }
    [data-testid="stAlert"][data-baseweb*="positive"] *,
    .stSuccess * { color: #065F46 !important; }
    
    /* Info alerts */
    [data-testid="stAlert"][data-baseweb*="info"],
    div[data-baseweb="notification"][kind="info"],
    .element-container .stInfo {
        background-color: #DBEAFE !important;
    }
    [data-testid="stAlert"][data-baseweb*="info"] *,
    .stInfo * { color: #1E3A8A !important; }
    
    /* Warning alerts */
    [data-testid="stAlert"][data-baseweb*="warning"],
    .stWarning {
        background-color: #FEF3C7 !important;
    }
    .stWarning * { color: #92400E !important; }
    
    /* Error alerts */
    [data-testid="stAlert"][data-baseweb*="negative"],
    .stError {
        background-color: #FEE2E2 !important;
    }
    .stError * { color: #991B1B !important; }
    
    /* ═══════════════════════════════════════════════════════════
       BADGES
       ═══════════════════════════════════════════════════════════ */
    .status-badge {
        display: inline-block;
        padding: 0.4rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.875rem;
    }
    .badge-success { background: #D1FAE5; color: #065F46 !important; }
    .badge-warning { background: #FEF3C7; color: #92400E !important; }
    .badge-error   { background: #FEE2E2; color: #991B1B !important; }
    .badge-info    { background: #DBEAFE; color: #1E40AF !important; }
    
    /* ═══════════════════════════════════════════════════════════
       TIPOGRAFÍA
       ═══════════════════════════════════════════════════════════ */
    h1 { color: #0F172A !important; font-weight: 800 !important; }
    h2 { color: #0F172A !important; font-weight: 700 !important; margin-top: 2rem !important; }
    h3 { color: #1E3A8A !important; font-weight: 600 !important; }
    h4 { color: #0F172A !important; font-weight: 600 !important; }
    h5, h6 { color: #334155 !important; font-weight: 600 !important; }
    
    p, li, span, div, label { color: #0F172A !important; }
    small, .stCaption { color: #334155 !important; font-size: 0.875rem !important; }
    .stMarkdown, .stMarkdown p, .stMarkdown div, .stMarkdown span { color: #0F172A !important; }
    strong, b { color: #1E3A8A !important; font-weight: 700 !important; }
    
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, #CBD5E1, transparent);
        margin: 2rem 0;
    }
    
    /* Spinner */
    .stSpinner > div { border-top-color: #3B82F6 !important; }
    .stSpinner > div + div { color: #334155 !important; }
    
    /* ═══════════════════════════════════════════════════════════
       TABLAS
       ═══════════════════════════════════════════════════════════ */
    .dataframe { border-radius: 8px; overflow: hidden; }
    .dataframe th { background-color: #1E3A8A !important; color: #FFFFFF !important; font-weight: 600 !important; }
    .dataframe td { color: #0F172A !important; background-color: #FFFFFF !important; }
    .dataframe tr:nth-child(even) td { background-color: #F8FAFC !important; }
    
    /* Container */
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1400px; }
    
    /* Info boxes */
    .success-box { background: #D1FAE5; border-left: 4px solid #059669; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
    .success-box * { color: #065F46 !important; }
    .warning-box { background: #FEF3C7; border-left: 4px solid #F59E0B; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
    .warning-box * { color: #92400E !important; }
    .info-box { background: #DBEAFE; border-left: 4px solid #3B82F6; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
    .info-box *, .info-box strong, .info-box span, .info-box p { color: #1E3A8A !important; }
    
    /* Inputs */
    input, textarea, select { color: #0F172A !important; background-color: #FFFFFF !important; }
    
    /* Responsive */
    @media (max-width: 768px) {
        .app-title { font-size: 1.8rem; }
        .score-circle { width: 80px; height: 80px; font-size: 1.8rem; }
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

def find_column(df, variants):
    """
    Busca una columna en el DataFrame probando múltiples variantes
    de nombre (case-insensitive, con/sin tildes, singular/plural).
    Retorna el nombre real de la columna o None.
    """
    df_cols_lower = {col.strip().lower(): col for col in df.columns}
    for variant in variants:
        key = variant.strip().lower()
        if key in df_cols_lower:
            return df_cols_lower[key]
    return None


def build_column_map(df):
    """
    Construye un diccionario que mapea nombres lógicos a nombres reales
    de columnas del DataFrame, probando múltiples variantes comunes.
    """
    mappings = {
        'nombre': [
            'Nombre', 'Nombres', 'NOMBRE', 'NOMBRES',
            'nombre', 'nombres', 'Name', 'Primer Nombre',
            'nombre completo', 'Nombre Completo', 'NOMBRE COMPLETO'
        ],
        'apellidos': [
            'Apellidos', 'Apellido', 'APELLIDOS', 'APELLIDO',
            'apellidos', 'apellido', 'Last Name', 'Surname',
            'Apellido Paterno', 'apellido paterno'
        ],
        'correo': [
            'Correo electrónico', 'Correo Electrónico', 'CORREO ELECTRÓNICO',
            'Correo electronico', 'Correo Electronico', 'CORREO ELECTRONICO',
            'correo electrónico', 'correo electronico',
            'Correo', 'correo', 'CORREO',
            'Email', 'email', 'EMAIL', 'E-mail', 'e-mail',
            'Mail', 'mail', 'Dirección de correo', 'Direccion de correo'
        ],
        'edad': [
            'Edad', 'edad', 'EDAD', 'Age', 'age'
        ],
        'programa': [
            'Programa', 'programa', 'PROGRAMA',
            'Carrera', 'carrera', 'CARRERA',
            'Programa Académico', 'Programa Academico',
            'programa académico', 'programa academico',
            'Especialidad', 'especialidad', 'ESPECIALIDAD',
            'Facultad', 'facultad'
        ],
        'respuesta_1': [
            'Respuesta 1', 'respuesta 1', 'RESPUESTA 1',
            'Respuesta1', 'respuesta1', 'R1', 'r1',
            'Pregunta 1', 'pregunta 1', 'P1', 'p1'
        ],
        'respuesta_2': [
            'Respuesta 2', 'respuesta 2', 'RESPUESTA 2',
            'Respuesta2', 'respuesta2', 'R2', 'r2',
            'Pregunta 2', 'pregunta 2', 'P2', 'p2'
        ],
        'respuesta_3': [
            'Respuesta 3', 'respuesta 3', 'RESPUESTA 3',
            'Respuesta3', 'respuesta3', 'R3', 'r3',
            'Pregunta 3', 'pregunta 3', 'P3', 'p3'
        ],
    }
    
    col_map = {}
    for logical_name, variants in mappings.items():
        found = find_column(df, variants)
        col_map[logical_name] = found  # None si no se encontró
    
    return col_map


def safe_get(row, col_name, default='N/A'):
    """Obtiene valor de una fila de forma segura."""
    if col_name is None:
        return default
    val = row.get(col_name, default)
    if pd.isna(val) or str(val).strip() == '':
        return default
    return str(val).strip()


# ═══════════════════════════════════════════════════════════════════════
# FUNCIONES DE EXTRACCIÓN DE TEXTO
# ═══════════════════════════════════════════════════════════════════════

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


# ═══════════════════════════════════════════════════════════════════════
# FUNCIÓN DE ANÁLISIS CON OPENAI
# ═══════════════════════════════════════════════════════════════════════

def analyze_admission_form(text_content, retry_count=0):
    system_prompt = """ROL Y CONTEXTO:
Actúa como experto en Psicología Educativa y Motivación, especializado en la Teoría de la Autodeterminación (Self-Determination Theory, SDT) de Ryan y Deci.

PROPÓSITO:
Identificar y caracterizar perfiles motivacionales de estudiantes en relación con su elección de carrera, experiencias formativas y proyección del uso del aprendizaje.

CONTEXTO INSTITUCIONAL:
Universidad Continental - Modalidad a Distancia
Población diversa: todo Perú, 18+ años, muchos trabajan y tienen familia
Objetivo: diagnosticar niveles motivacionales, NO filtrar ni evaluar ortografía

MARCO TEÓRICO - SDT:
El continuo motivacional es jerárquico y no acumulativo. La regulación menos autónoma presente limita el nivel funcional de autodeterminación.

ESCALA DE EVALUACIÓN (1-6):
6 = Motivación Intrínseca: interés genuino, disfrute, curiosidad inherente
5 = Regulación Integrada: coherencia con identidad, valores centrales, proyecto de vida
4 = Regulación Identificada: utilidad personal significativa, metas importantes
3 = Regulación Introyectada: presión interna, culpa, orgullo, autoexigencia
2 = Regulación Externa: recompensas/presiones externas, demandas sociales
1 = Amotivación: sin razón clara, desinterés, resignación

RÚBRICA DETALLADA POR NIVEL:

NIVEL 6 - MOTIVACIÓN INTRÍNSECA:
Criterio general: Interés genuino, disfrute, curiosidad o satisfacción personal.
- Elección de carrera: Se basa en el agrado inherente por los contenidos o actividades propias de la carrera
- Experiencia: Disfrute del proceso, interés espontáneo, sensación de flujo
- Uso futuro: Desea aplicar lo aprendido por interés y disfrute personal
Indicadores: "me gusta", "lo disfruto", "me interesa mucho", "me apasiona"
NO asignar si: el interés se justifica por utilidad, resultados, metas, identidad o impacto social

NIVEL 5 - REGULACIÓN INTEGRADA:
Criterio general: Coherencia con identidad, valores centrales y proyecto de vida.
Indicadores: "es coherente con mis valores", "encaja con mi forma de desarrollarme", "es parte de mi proyecto"
NO asignar si: solo hay disfrute (→6), solo utilidad (→4)

NIVEL 4 - REGULACIÓN IDENTIFICADA:
Criterio general: Reconoce importancia y utilidad personal para metas significativas.
Indicadores: "es importante para mí", "me permite desarrollarme", "me ayuda a lograr mis metas"
NO asignar si: solo hay recompensas externas (→2), se menciona identidad/proyecto vital (→5)

NIVEL 3 - REGULACIÓN INTROYECTADA:
Criterio general: Presión interna, necesidad de validación, evitar emociones negativas.
Indicadores: "sentía que debía", "no quería fallar", "quería demostrar", "me sentiría mal si no"
NO asignar si: hay demandas externas explícitas (→2), hay valor personal o metas (→4)

NIVEL 2 - REGULACIÓN EXTERNA:
Criterio general: Recompensas externas, demandas sociales, control externo.
Indicadores: "tiene buena salida laboral", "da estabilidad", "mis padres querían", "para conseguir trabajo"
NO asignar si: hay culpa/orgullo (→3), valor personal explícito (→4)

NIVEL 1 - AMOTIVACIÓN:
Criterio general: Incapaz de dar razón clara, desinterés, falta de control.
Indicadores: "no lo tengo claro", "no sé por qué", "me da igual", "me obligaron"
SE ASIGNA si ocurre AL MENOS UNA condición de amotivación

REGLAS DE ASIGNACIÓN (OBLIGATORIAS):
1. Solo asignar niveles con indicadores EXPLÍCITOS en el texto
2. Cada respuesta se clasifica en UN solo nivel por criterio
3. Si coexisten indicadores de varios niveles → asignar el nivel INFERIOR
4. Cada pregunta (P1, P2, P3) se evalúa de manera INDEPENDIENTE
5. PERFIL FINAL: min(P1, P2, P3) con excepción 2-de-3
   Excepción: Si dos coinciden y tercera está 1 nivel abajo → perfil = nivel coincidente
   Si alguna = 1 → perfil final máximo = 2

ESTRUCTURA JSON:
{
  "informacion_extraida": {
    "nombre": "...",
    "apellidos": "...",
    "edad": "...",
    "programa": "...",
    "correo": "..."
  },
  "evaluacion_motivacional": {
    "eleccion_carrera": {
      "puntaje": 1-6,
      "tipo_motivacion": "...",
      "justificacion": "Evidencia textual breve"
    },
    "experiencia_relacionada": {...},
    "uso_futuro": {...}
  },
  "necesidades_psicologicas": {
    "autonomia": "Alta/Media/Baja - Análisis",
    "competencia": "Alta/Media/Baja - Análisis",
    "relacion": "Alta/Media/Baja - Análisis"
  },
  "calificacion_real": 14,
  "calificacion_sobre_20": 15.56,
  "perfil_motivacional_final": "Identificado",
  "regla_aplicada": "min(4,5,4)=4",
  "recomendaciones": "...",
  "nivel_motivacional_general": "Predominantemente Identificado"
}

CÁLCULOS:
- calificacion_real = P1+P2+P3 (máx 18)
- calificacion_sobre_20 = (real/18)*20 (2 decimales)

IMPORTANTE: Solo JSON, sin markdown, comillas dobles, evidencia explícita"""

    user_prompt = f"""PREGUNTAS DEL FORMULARIO:
1. ¿Qué características de esta carrera llamaron tu atención y cuál es la razón principal por la que decidiste postular a ella?
2. Relata una experiencia donde hayas puesto en práctica habilidades relacionadas con esta carrera. Describe cómo te sentiste mientras realizabas dicha actividad y qué descubriste de tu vocación profesional.
3. Imagina que ya terminaste tus estudios. ¿Cómo aplicarías lo aprendido en tu formación profesional y qué impactos te gustaría lograr?

FORMULARIO:
{text_content}

PROCESO: 1)Leer 2)Identificar indicadores 3)Verificar condiciones 4)Si coexisten→inferior 5)Evaluar independiente 6)Calcular perfil 7)JSON

Responde ÚNICAMENTE con JSON válido (sin markdown)"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.2,
            max_tokens=3500
        )
        
        content = response.choices[0].message.content.strip()
        
        # Limpiar markdown
        if content.startswith('```'):
            lines = content.split('\n')
            content = '\n'.join(lines[1:-1]) if len(lines) > 2 else content
            content = content.replace('```json', '').replace('```', '').strip()
        
        result = json.loads(content)
        result['tokens_used'] = response.usage.total_tokens
        result['timestamp'] = datetime.now().isoformat()
        result['success'] = True
        
        return result
        
    except json.JSONDecodeError as e:
        if retry_count < 2:
            import time
            time.sleep(1)
            return analyze_admission_form(text_content, retry_count + 1)
        return {
            "success": False, 
            "error": f"Error al parsear JSON después de {retry_count + 1} intentos",
            "detail": str(e)
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ═══════════════════════════════════════════════════════════════════════
# PROCESAR REGISTROS EXCEL (con mapeo flexible de columnas)
# ═══════════════════════════════════════════════════════════════════════

def process_excel_records(df, progress_bar, status_text):
    results = []
    total = len(df)
    
    # Construir mapeo flexible de columnas
    col_map = build_column_map(df)
    
    # Mostrar qué columnas se detectaron
    detected = {k: v for k, v in col_map.items() if v is not None}
    missing = {k for k, v in col_map.items() if v is None}
    
    if missing:
        st.warning(
            f"⚠️ Columnas no encontradas: **{', '.join(missing)}**. "
            f"Se usará 'N/A' para esos campos.\n\n"
            f"**Columnas detectadas en tu archivo:** {', '.join(df.columns.tolist())}"
        )
    
    for idx, row in df.iterrows():
        status_text.markdown(f"**Procesando:** Registro {idx + 1} de {total}")
        progress_bar.progress((idx + 1) / total)
        
        # Extraer valores con mapeo flexible
        nombre = safe_get(row, col_map['nombre'])
        apellidos = safe_get(row, col_map['apellidos'])
        correo = safe_get(row, col_map['correo'])
        edad = safe_get(row, col_map['edad'])
        programa = safe_get(row, col_map['programa'])
        resp1 = safe_get(row, col_map['respuesta_1'], 'Sin respuesta')
        resp2 = safe_get(row, col_map['respuesta_2'], 'Sin respuesta')
        resp3 = safe_get(row, col_map['respuesta_3'], 'Sin respuesta')
        
        form_text = f"""
Nombre: {nombre}
Apellidos: {apellidos}
Correo: {correo}
Edad: {edad}
Programa: {programa}

Pregunta 1: ¿Por qué elegiste esta carrera?
Respuesta 1: {resp1}

Pregunta 2: ¿Qué experiencia tienes relacionada con esta carrera?
Respuesta 2: {resp2}

Pregunta 3: ¿Cómo planeas usar lo que aprendas?
Respuesta 3: {resp3}
"""
        
        # Verificar respuestas vacías
        missing_responses = []
        if resp1 == 'Sin respuesta':
            missing_responses.append('Respuesta 1')
        if resp2 == 'Sin respuesta':
            missing_responses.append('Respuesta 2')
        if resp3 == 'Sin respuesta':
            missing_responses.append('Respuesta 3')
        
        if missing_responses:
            results.append({
                'success': False,
                'registro_numero': idx + 1,
                'nombre': nombre,
                'apellidos': apellidos,
                'correo': correo,
                'error': f"Campos faltantes: {', '.join(missing_responses)}"
            })
            continue
        
        analysis = analyze_admission_form(form_text)
        
        result = {
            'registro_numero': idx + 1,
            'nombre': nombre,
            'apellidos': apellidos,
            'correo': correo,
            'success': analysis.get('success', False),
        }
        
        if analysis.get('success'):
            result['analysis'] = analysis
        else:
            result['error'] = analysis.get('error', 'Error desconocido')
        
        results.append(result)
    
    return results


# ═══════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE RESULTADOS
# ═══════════════════════════════════════════════════════════════════════

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
        'N°', 'Nombre', 'Apellidos', 'Correo', 'Calif. Real', 'Calif. /20',
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
            r.get('nombre', ''),
            r.get('apellidos', ''),
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


# ═══════════════════════════════════════════════════════════════════════
# INTERFAZ PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════

def main():
    # Header profesional
    st.markdown("""
    <div class="app-header">
        <h1 class="app-title">🎓 Sistema de Análisis de Admisión</h1>
        <p class="app-subtitle">Análisis Motivacional basado en la Teoría de la Autodeterminación (SDT)</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Información de uso
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='background: #FFFFFF; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;'>
            <p style='margin: 0; color: #334155; font-size: 0.95rem;'>
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
        
        # Info del archivo
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"""
            <div class='info-box'>
                📁 <strong>{uploaded_file.name}</strong>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class='info-box'>
                💾 {file_size:.2f} MB
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
                    # ═══ PROCESAMIENTO MASIVO ═══
                    df = read_excel_file(uploaded_file)
                    
                    if df is not None:
                        st.success(f"✅ {len(df)} registros detectados")
                        
                        # Mostrar columnas encontradas
                        col_map = build_column_map(df)
                        with st.expander("🔍 Mapeo de columnas detectado", expanded=False):
                            for logical, real in col_map.items():
                                status = f"✅ `{real}`" if real else "❌ No encontrada"
                                st.markdown(f"**{logical}** → {status}")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        results = process_excel_records(df, progress_bar, status_text)
                        
                        status_text.markdown("✅ **Análisis completado exitosamente**")
                        progress_bar.progress(1.0)
                        
                        st.session_state['batch_results'] = results
                        st.session_state['batch_filename'] = uploaded_file.name
                        
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
                            # Construir label legible
                            nombre_display = result.get('nombre', 'N/A')
                            apellidos_display = result.get('apellidos', 'N/A')
                            correo_display = result.get('correo', 'N/A')
                            
                            # Evitar mostrar N/A si no hay datos
                            label_parts = []
                            if apellidos_display != 'N/A':
                                label_parts.append(apellidos_display)
                            if nombre_display != 'N/A':
                                label_parts.append(nombre_display)
                            
                            if label_parts:
                                name_label = ", ".join(label_parts)
                            else:
                                name_label = f"Registro {result['registro_numero']}"
                            
                            extra_info = f" • {correo_display}" if correo_display != 'N/A' else ""
                            
                            with st.expander(
                                f"**{result['registro_numero']}. {name_label}**{extra_info}",
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
                                        st.markdown(f"**📊 Calificación Real:** {analysis.get('calificacion_real', 'N/A')}/18")
                                        nivel = analysis.get('nivel_motivacional_general', 'N/A')
                                        st.markdown(f"**🎯 Nivel Motivacional:** {nivel}")
                                    
                                    st.markdown("---")
                                    
                                    # Evaluación motivacional
                                    st.markdown("#### 📝 Evaluación Motivacional Detallada")
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
                                        st.markdown("#### 🧠 Necesidades Psicológicas (SDT)")
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
                                        st.markdown("#### 💡 Recomendaciones Pedagógicas")
                                        st.info(analysis['recomendaciones'])
                                else:
                                    st.error(f"❌ **Error:** {result.get('error', 'Error desconocido')}")
                
                else:
                    # ═══ PROCESAMIENTO INDIVIDUAL ═══
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
                                    <p style='color: #64748B; margin-top: 1rem; font-weight: 600;'>Calificación Final</p>
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
                                        st.markdown(f"**Justificación:** {item.get('justificacion')}")
                            
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
    
    # Footer informativo
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; padding: 2rem; color: #94A3B8; font-size: 0.85rem;'>
        <p>Dirección de Gestión de la Información • Universidad Continental</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

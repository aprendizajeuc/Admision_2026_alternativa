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
    page_title="Sistema de Analisis de Admision - SDT",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    :root {
        --primary-blue: #1E3A8A;
        --secondary-blue: #3B82F6;
        --success-green: #059669;
        --warning-orange: #F59E0B;
        --error-red: #DC2626;
        --bg-light: #F8FAFC;
        --text-dark: #0F172A;
        --text-medium: #334155;
        --text-light: #64748B;
        --border-color: #E2E8F0;
    }
    .main, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"],
    [data-testid="stMainBlockContainer"], [data-testid="stVerticalBlock"],
    [data-testid="stAppViewBlockContainer"], section[data-testid="stSidebar"],
    .block-container {
        background-color: #F1F5F9 !important; color: #0F172A !important;
    }
    .main { background: linear-gradient(180deg, #EFF6FF 0%, #F1F5F9 100%) !important; padding: 2rem 1rem; }
    div[data-testid="stFileUploader"] { background: #FFFFFF !important; border-radius: 12px; padding: 2rem; border: 2px dashed #3B82F6 !important; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    div[data-testid="stFileUploader"]:hover { border-color: #1E3A8A !important; }
    div[data-testid="stFileUploader"] label, div[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] span, div[data-testid="stFileUploader"] p,
    div[data-testid="stFileUploader"] div { color: #0F172A !important; font-weight: 500 !important; }
    div[data-testid="stFileUploader"] section, div[data-testid="stFileUploader"] section > div,
    [data-testid="stFileUploaderDropzone"] { background-color: #FFFFFF !important; background: #FFFFFF !important; color: #0F172A !important; border-color: #3B82F6 !important; }
    div[data-testid="stFileUploader"] button, [data-testid="stFileUploaderDropzone"] button { background-color: #3B82F6 !important; color: #FFFFFF !important; border: none !important; font-weight: 600 !important; }
    div[data-testid="stFileUploader"] button:hover { background-color: #1E3A8A !important; }
    [data-testid="stFileUploaderDropzoneInstructions"] div, [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small { color: #334155 !important; }
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] * { background-color: #F8FAFC !important; color: #0F172A !important; }
    .app-header { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); padding: 2rem; border-radius: 16px; margin-bottom: 2rem; box-shadow: 0 10px 30px rgba(30,58,138,0.2); text-align: center; }
    .app-title { font-size: 2.5rem; font-weight: 800; color: #FFFFFF !important; margin: 0; }
    .app-subtitle { font-size: 1.1rem; color: #F1F5F9 !important; margin-top: 0.5rem; }
    .stButton > button { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); color: #FFFFFF !important; border: none; padding: 0.75rem 2rem; font-size: 1rem; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(30,58,138,0.3); }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(30,58,138,0.4); }
    div[data-testid="metric-container"] { background: #FFFFFF !important; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid #3B82F6; }
    div[data-testid="metric-container"] label { color: #334155 !important; font-size: 0.875rem; font-weight: 600; }
    div[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #1E3A8A !important; font-size: 2rem; font-weight: 700; }
    .score-circle { display: inline-flex; width: 100px; height: 100px; border-radius: 50%; background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); color: #FFFFFF !important; align-items: center; justify-content: center; font-size: 2.2rem; font-weight: 800; box-shadow: 0 8px 20px rgba(30,58,138,0.3); position: relative; }
    .score-circle::after { content: '/20'; position: absolute; bottom: 10px; right: 10px; font-size: 0.7rem; opacity: 0.9; color: #FFFFFF !important; }
    .stProgress > div > div > div > div { background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%); }
    div[data-testid="stExpander"] { background-color: #FFFFFF !important; border: 1px solid #E2E8F0 !important; border-radius: 8px !important; margin-bottom: 0.5rem; overflow: hidden; }
    div[data-testid="stExpander"] details > summary { background-color: #FFFFFF !important; color: #0F172A !important; font-weight: 600; padding: 0.75rem 1rem; border: none !important; }
    div[data-testid="stExpander"] details > summary:hover { background-color: #F8FAFC !important; }
    div[data-testid="stExpander"] details[open] > summary { background-color: #FFFFFF !important; color: #0F172A !important; border-bottom: 2px solid #3B82F6 !important; }
    div[data-testid="stExpander"] details > summary *, div[data-testid="stExpander"] summary [data-testid="stMarkdownContainer"] * { color: #0F172A !important; }
    div[data-testid="stExpander"] details > div, div[data-testid="stExpander"] details[open] > div,
    div[data-testid="stExpander"] details[open] > div > div, [data-testid="stExpanderDetails"] { background-color: #FFFFFF !important; background: #FFFFFF !important; color: #0F172A !important; }
    div[data-testid="stExpander"] details[open] p, div[data-testid="stExpander"] details[open] div,
    div[data-testid="stExpander"] details[open] span, div[data-testid="stExpander"] details[open] strong,
    [data-testid="stExpanderDetails"] *, [data-testid="stExpanderDetails"] p,
    [data-testid="stExpanderDetails"] div, [data-testid="stExpanderDetails"] span { color: #0F172A !important; background-color: transparent !important; }
    div[data-testid="stExpander"] [data-testid="stAlert"], [data-testid="stExpanderDetails"] [data-testid="stAlert"] { background-color: #DBEAFE !important; }
    div[data-testid="stExpander"] [data-testid="stAlert"] *, [data-testid="stExpanderDetails"] [data-testid="stAlert"] * { color: #1E3A8A !important; }
    .stDownloadButton > button { background: linear-gradient(135deg, #059669 0%, #10B981 100%); color: #FFFFFF !important; border: none; padding: 0.75rem 2rem; font-size: 1rem; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 12px rgba(5,150,105,0.3); }
    .stDownloadButton > button:hover { transform: translateY(-2px); }
    .stAlert { border-radius: 8px; border-left: 4px solid; }
    [data-testid="stAlert"][data-baseweb*="positive"], .stSuccess { background-color: #D1FAE5 !important; }
    [data-testid="stAlert"][data-baseweb*="positive"] *, .stSuccess * { color: #065F46 !important; }
    [data-testid="stAlert"][data-baseweb*="info"], .stInfo { background-color: #DBEAFE !important; }
    [data-testid="stAlert"][data-baseweb*="info"] *, .stInfo * { color: #1E3A8A !important; }
    [data-testid="stAlert"][data-baseweb*="warning"], .stWarning { background-color: #FEF3C7 !important; }
    .stWarning * { color: #92400E !important; }
    [data-testid="stAlert"][data-baseweb*="negative"], .stError { background-color: #FEE2E2 !important; }
    .stError * { color: #991B1B !important; }
    .status-badge { display: inline-block; padding: 0.4rem 1rem; border-radius: 20px; font-weight: 600; font-size: 0.875rem; }
    .badge-success { background: #D1FAE5; color: #065F46 !important; }
    .badge-warning { background: #FEF3C7; color: #92400E !important; }
    .badge-error { background: #FEE2E2; color: #991B1B !important; }
    .badge-info { background: #DBEAFE; color: #1E40AF !important; }
    h1 { color: #0F172A !important; font-weight: 800 !important; }
    h2 { color: #0F172A !important; font-weight: 700 !important; margin-top: 2rem !important; }
    h3 { color: #1E3A8A !important; font-weight: 600 !important; }
    h4 { color: #0F172A !important; font-weight: 600 !important; }
    p, li, span, div, label { color: #0F172A !important; }
    small, .stCaption { color: #334155 !important; }
    .stMarkdown, .stMarkdown p, .stMarkdown div, .stMarkdown span { color: #0F172A !important; }
    strong, b { color: #1E3A8A !important; font-weight: 700 !important; }
    hr { border: none; height: 2px; background: linear-gradient(90deg, transparent, #CBD5E1, transparent); margin: 2rem 0; }
    .dataframe th { background-color: #1E3A8A !important; color: #FFFFFF !important; font-weight: 600 !important; }
    .dataframe td { color: #0F172A !important; background-color: #FFFFFF !important; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1400px; }
    .info-box { background: #DBEAFE; border-left: 4px solid #3B82F6; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
    .info-box *, .info-box strong, .info-box span, .info-box p { color: #1E3A8A !important; }
    input, textarea, select { color: #0F172A !important; background-color: #FFFFFF !important; }
    @media (max-width: 768px) { .app-title { font-size: 1.8rem; } .score-circle { width: 80px; height: 80px; font-size: 1.8rem; } }
</style>
""", unsafe_allow_html=True)

@st.cache_resource
def get_openai_client():
    api_key = None
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except:
        api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        st.error("No se encontro la clave API de OpenAI")
        st.info("**Para uso local:** Configura OPENAI_API_KEY en .streamlit/secrets.toml")
        st.info("**Para Streamlit Cloud:** Configura OPENAI_API_KEY en Settings -> Secrets")
        st.stop()
    return OpenAI(api_key=api_key)

client = get_openai_client()

def find_column(df, variants):
    df_cols_lower = {col.strip().lower(): col for col in df.columns}
    for variant in variants:
        key = variant.strip().lower()
        if key in df_cols_lower:
            return df_cols_lower[key]
    return None

def build_column_map(df):
    mappings = {
        'nombre': ['Nombre', 'Nombres', 'NOMBRE', 'NOMBRES', 'nombre', 'nombres', 'Name', 'Primer Nombre', 'nombre completo', 'Nombre Completo'],
        'apellidos': ['Apellidos', 'Apellido', 'APELLIDOS', 'APELLIDO', 'apellidos', 'apellido', 'Last Name', 'Surname', 'Apellido Paterno'],
        'correo': ['Correo electronico', 'Correo Electronico', 'correo electronico', 'Correo', 'correo', 'CORREO', 'Email', 'email', 'EMAIL', 'E-mail', 'Mail', 'mail', 'Direccion de correo'],
        'edad': ['Edad', 'edad', 'EDAD', 'Age'],
        'programa': ['Programa', 'programa', 'PROGRAMA', 'Carrera', 'carrera', 'CARRERA', 'Programa Academico', 'Especialidad', 'especialidad', 'Facultad'],
        'respuesta_1': ['Respuesta 1', 'respuesta 1', 'RESPUESTA 1', 'Respuesta1', 'R1', 'r1', 'Pregunta 1', 'P1'],
        'respuesta_2': ['Respuesta 2', 'respuesta 2', 'RESPUESTA 2', 'Respuesta2', 'R2', 'r2', 'Pregunta 2', 'P2'],
        'respuesta_3': ['Respuesta 3', 'respuesta 3', 'RESPUESTA 3', 'Respuesta3', 'R3', 'r3', 'Pregunta 3', 'P3'],
    }
    col_map = {}
    for logical_name, variants in mappings.items():
        col_map[logical_name] = find_column(df, variants)
    return col_map

def safe_get(row, col_name, default='N/A'):
    if col_name is None:
        return default
    val = row.get(col_name, default)
    if pd.isna(val) or str(val).strip() == '':
        return default
    return str(val).strip()

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
        return docx2txt.process(file)
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
        ext = file.name.split('.')[-1].lower()
        return pd.read_csv(file) if ext == 'csv' else pd.read_excel(file)
    except Exception as e:
        st.error(f"Error al leer archivo: {str(e)}")
        return None


# =====================================================================
# PROMPT v3 - Sin frases literales + reglas satisfaccion/recompensa
# =====================================================================

SYSTEM_PROMPT = """ROL: Experto en Psicologia Educativa especializado en Teoria de la Autodeterminacion (SDT) de Ryan y Deci.

CONTEXTO: Universidad Continental Peru, modalidad a distancia. Poblacion diversa, 18+ anios. Objetivo: diagnosticar motivacion, NO evaluar ortografia ni redaccion.

ESCALA SDT (1-6):
6=Intrinseca | 5=Integrada | 4=Identificada | 3=Introyectada | 2=Externa | 1=Amotivacion

=====================================================================
RUBRICA CON CRITERIOS GENERALES
=====================================================================

NIVEL 6 - MOTIVACION INTRINSECA
Definicion: Disfrute genuino del PROCESO, curiosidad inherente por la actividad misma.
El foco DEBE estar en el proceso o actividad, NO en resultados, impactos, identidad ni metas.
Indicadores tipicos: expresiones de gusto, disfrute, curiosidad, fascinacion, estimulacion intelectual referidas a la actividad en si.
NO asignar si: el interes se justifica por utilidad, impacto social, metas, identidad o resultados tangibles.

NIVEL 5 - REGULACION INTEGRADA
Definicion: La actividad ES PARTE de la identidad y proyecto de vida de la persona. Hay coherencia profunda con valores centrales.
Debe haber declaracion EXPLICITA de que la actividad forma parte de quien la persona ES o quiere SER.
Indicadores tipicos: expresiones de coherencia con valores, sentido de identidad profesional, integracion con proyecto de vida, declaraciones de que la actividad define quien es.
NO asignar si: solo hay disfrute sin mencion de identidad (->6), o solo hay utilidad/metas (->4).

NIVEL 4 - REGULACION IDENTIFICADA
Definicion: Reconoce VALOR e IMPORTANCIA personal. Elige porque ve utilidad para metas significativas o desarrollo propio.
La persona ELIGE libremente porque reconoce valor, NO actua por obligacion emocional ni por recompensa externa.
Indicadores tipicos: expresiones de importancia para el desarrollo, valoracion del aprendizaje, reconocimiento de utilidad para metas personales, deseo de contribuir o aportar.
NO asignar si: el motor principal es una recompensa tangible (->2), o una obligacion emocional interna (->3).

NIVEL 3 - REGULACION INTROYECTADA
Definicion: PRESION EMOCIONAL INTERNA. Actua para evitar emociones negativas o buscar validacion emocional del yo.
El motor es una EMOCION dirigida al yo (verguenza, culpa, orgullo, miedo a decepcionar, necesidad de demostrar capacidad), NO un resultado tangible externo.
Indicadores tipicos: expresiones de verguenza, culpa, necesidad de demostrar, miedo a fallar, deseo de que otros se sientan orgullosos, necesidad de sentir que se cumple, autoexigencia emocional.
NO asignar si: el motor principal es una recompensa tangible externa (->2), o hay valor personal genuino hacia el aprendizaje (->4).
REGLA ESPECIAL - SATISFACCION Y RECOMPENSA EXTERNA: No se clasificara como regulacion introyectada cuando el texto exprese unicamente la busqueda de una recompensa externa (nota, reconocimiento, premio), aun cuando el postulante mencione satisfaccion posterior por haberla obtenido. La satisfaccion derivada de una recompensa externa NO constituye presion interna. Si alguien dice que busco una buena nota y se sintio satisfecho al obtenerla, eso sigue siendo regulacion externa (2), no introyectada (3).

NIVEL 2 - REGULACION EXTERNA
Definicion: RECOMPENSAS O PRESIONES TANGIBLES del exterior. Busca resultados concretos y medibles.
Lo que motiva es algo TANGIBLE: dinero, empleo, notas, estatus social, estabilidad economica, reconocimiento formal, prestigio, beneficios laborales.
Indicadores tipicos: expresiones sobre salario, empleabilidad, estabilidad, notas, reconocimiento, prestigio profesional, demanda laboral, beneficios materiales.
NO asignar si: hay emociones internas como motor principal (verguenza, culpa, orgullo, complacer) (->3), o hay valor personal genuino (->4).
REGLA ESPECIAL - RECOMPENSA EXTERNA CON EMOCION POSITIVA: Cuando una respuesta contenga simultaneamente una recompensa externa explicita Y una emocion positiva posterior (por ejemplo, satisfaccion por obtener una nota o reconocimiento), se asignara regulacion externa (2), SALVO que exista mencion explicita de culpa, verguenza, orgullo como motor, complacer a otros o autoexigencia. La secuencia "busque X tangible y me senti bien al lograrlo" es regulacion externa, no introyectada.

NIVEL 1 - AMOTIVACION
Definicion: Sin razon clara, desinteres, inercia, resignacion, falta de control percibido.
Indicadores tipicos: ausencia de razon, expresiones de indiferencia, falta de interes, inercia, resignacion, percepcion de falta de control sobre la decision, incertidumbre total sobre el futuro profesional.
Regla: Basta UNA senal clara de amotivacion para asignar este nivel.

=====================================================================
REGLAS DE DISCRIMINACION OBLIGATORIAS (D1-D4)
=====================================================================

Antes de asignar CUALQUIER puntaje, aplica la regla de discriminacion relevante:

--- REGLA D1: Discriminar NIVEL 2 vs NIVEL 3 ---
Pregunta clave: "Que busca la persona: un RESULTADO TANGIBLE o resolver/gestionar una EMOCION?"

TANGIBLE (dinero, empleo, notas, estatus, reconocimiento, estabilidad) = NIVEL 2
EMOCION (verguenza, culpa, orgullo como motor, miedo a decepcionar, complacer) = NIVEL 3

Principios de discriminacion:
- Lo que define el nivel NO es quien aparece en la frase (padres, amigos, jefe), sino QUE TIPO DE COSA motiva la accion.
- Si el texto menciona estatus social, respeto ajeno y dinero como motivadores = tangibles = NIVEL 2
- Si el texto menciona verguenza ante otros, evitar decepcionar, buscar que se sientan orgullosos = emociones = NIVEL 3
- Si el texto menciona buscar buenas notas o reconocimiento y luego sentirse satisfecho = la satisfaccion es consecuencia de lo tangible = NIVEL 2
- Si el texto menciona elegir algo para que otra persona sea feliz o no sufra = complacer emocionalmente = NIVEL 3
- Si el texto menciona elegir algo por imagen o presencia ante otros motivado por verguenza = emocion interna = NIVEL 3

--- REGLA D2: Discriminar NIVEL 3 vs NIVEL 4 ---
Pregunta clave: "La persona VALORA el aprendizaje/actividad, o actua por OBLIGACION EMOCIONAL?"

VALORA (reconoce importancia, ve utilidad, quiere contribuir, aprecia el aprendizaje) = NIVEL 4
OBLIGACION (sentir que cumple, demostrar capacidad, no fallar, no decepcionar, autoexigencia) = NIVEL 3

Principios de discriminacion:
- En nivel 4 la persona ELIGE porque ve valor en lo que aprende o en contribuir.
- En nivel 3 la persona se SIENTE OBLIGADA emocionalmente (aunque la obligacion venga de si misma).
- "Quiero contribuir a mejorar mi entorno" (elige por valor) = NIVEL 4
- "Necesito sentir que cumplo con lo esperado" (obligacion interna) = NIVEL 3
- "Valore las habilidades que desarrolle" (aprecia aprendizaje) = NIVEL 4
- "Me esforcé para no sentirme incompetente" (evitar emocion negativa) = NIVEL 3

--- REGLA D3: Discriminar NIVEL 4 vs NIVEL 5 ---
Pregunta clave: "Habla de METAS/UTILIDAD o de IDENTIDAD/PROYECTO DE VIDA?"

Metas, utilidad, desarrollo profesional, contribucion = NIVEL 4
Identidad, proyecto vital, coherencia con valores centrales, "quien quiero ser" = NIVEL 5

--- REGLA D4: Discriminar NIVEL 5 vs NIVEL 6 ---
Pregunta clave: "El foco esta en QUIEN SOY (identidad) o en lo que DISFRUTO HACER (proceso)?"

Coherencia con valores, identidad, proyecto de vida = NIVEL 5
Disfrute del proceso, curiosidad, placer en la actividad misma = NIVEL 6

=====================================================================
PROCESO DE EVALUACION (SEGUIR EN ORDEN ESTRICTO)
=====================================================================

Para CADA respuesta (P1, P2, P3):
1. Leer la respuesta completa sin prejuicios
2. Identificar TODOS los indicadores motivacionales presentes en el texto
3. Clasificar cada indicador segun la rubrica
4. Si hay indicadores de MULTIPLES niveles -> asignar el nivel MAS BAJO (menos autonomo)
5. VERIFICAR aplicando la regla D1, D2, D3 o D4 segun los niveles en juego
6. Si hay recompensa externa + satisfaccion posterior -> verificar REGLA ESPECIAL de nivel 2
7. Escribir justificacion citando la evidencia textual Y la regla aplicada

PERFIL FINAL:
- Base: Perfil = min(P1, P2, P3)
- Excepcion 2-de-3: Si dos puntajes coinciden y el tercero esta EXACTAMENTE 1 nivel abajo -> perfil = nivel coincidente
  Validos: (4,4,3)->4, (5,5,4)->5, (6,6,5)->6
  NO validos: (6,5,4)->4, (5,5,3)->3
- Seguridad: Si alguno = 1 -> perfil maximo = 2

CALCULOS:
- calificacion_real = P1+P2+P3 (max 18)
- calificacion_sobre_20 = (real/18)*20 con 2 decimales

=====================================================================
JSON DE RESPUESTA (responder SOLO con este JSON, sin markdown)
=====================================================================

{
  "informacion_extraida": {
    "nombre": "...", "apellidos": "...", "edad": "...", "programa": "...", "correo": "..."
  },
  "evaluacion_motivacional": {
    "eleccion_carrera": {
      "puntaje": 1-6,
      "tipo_motivacion": "nombre del nivel",
      "justificacion": "Evidencia: [cita textual]. Regla aplicada: [D1/D2/D3/D4 + explicacion]"
    },
    "experiencia_relacionada": {
      "puntaje": 1-6,
      "tipo_motivacion": "nombre del nivel",
      "justificacion": "Evidencia: [cita textual]. Regla aplicada: [D1/D2/D3/D4 + explicacion]"
    },
    "uso_futuro": {
      "puntaje": 1-6,
      "tipo_motivacion": "nombre del nivel",
      "justificacion": "Evidencia: [cita textual]. Regla aplicada: [D1/D2/D3/D4 + explicacion]"
    }
  },
  "necesidades_psicologicas": {
    "autonomia": "Alta/Media/Baja - Analisis breve",
    "competencia": "Alta/Media/Baja - Analisis breve",
    "relacion": "Alta/Media/Baja - Analisis breve"
  },
  "calificacion_real": 0,
  "calificacion_sobre_20": 0.00,
  "perfil_motivacional_final": "...",
  "regla_aplicada": "min(X,Y,Z)=N o excepcion aplicada",
  "recomendaciones": "...",
  "nivel_motivacional_general": "Predominantemente ..."
}"""

USER_PROMPT_TEMPLATE = """PREGUNTAS DEL FORMULARIO:
1. Que caracteristicas de esta carrera llamaron tu atencion y cual es la razon principal por la que decidiste postular a ella?
2. Relata una experiencia donde hayas puesto en practica habilidades relacionadas con esta carrera. Describe como te sentiste y que descubriste de tu vocacion.
3. Imagina que ya terminaste tus estudios. Como aplicarias lo aprendido y que impactos te gustaria lograr?

FORMULARIO:
{text_content}

INSTRUCCION FINAL:
1) Lee cada respuesta completa
2) Identifica todos los indicadores motivacionales
3) Si hay multiples niveles -> asigna el nivel inferior
4) APLICA la regla de discriminacion D1/D2/D3/D4 segun corresponda
5) Si hay recompensa externa + satisfaccion posterior -> verifica REGLA ESPECIAL nivel 2
6) Justifica con evidencia textual + regla aplicada
7) Calcula perfil final
8) Responde SOLO con JSON valido (sin markdown, sin backticks)"""


def analyze_admission_form(text_content, retry_count=0):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": USER_PROMPT_TEMPLATE.format(text_content=text_content)}
            ],
            temperature=0.1,
            max_tokens=3500
        )
        
        content = response.choices[0].message.content.strip()
        
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
        return {"success": False, "error": f"Error JSON tras {retry_count + 1} intentos", "detail": str(e)}
    except Exception as e:
        return {"success": False, "error": str(e)}


def process_excel_records(df, progress_bar, status_text):
    results = []
    total = len(df)
    col_map = build_column_map(df)
    
    missing = {k for k, v in col_map.items() if v is None}
    if missing:
        st.warning(f"Columnas no encontradas: **{', '.join(missing)}**. Se usara 'N/A'.\n\n**Columnas en archivo:** {', '.join(df.columns.tolist())}")
    
    for idx, row in df.iterrows():
        status_text.markdown(f"**Procesando:** Registro {idx + 1} de {total}")
        progress_bar.progress((idx + 1) / total)
        
        nombre = safe_get(row, col_map['nombre'])
        apellidos = safe_get(row, col_map['apellidos'])
        correo = safe_get(row, col_map['correo'])
        edad = safe_get(row, col_map['edad'])
        programa = safe_get(row, col_map['programa'])
        resp1 = safe_get(row, col_map['respuesta_1'], 'Sin respuesta')
        resp2 = safe_get(row, col_map['respuesta_2'], 'Sin respuesta')
        resp3 = safe_get(row, col_map['respuesta_3'], 'Sin respuesta')
        
        form_text = f"""Nombre: {nombre}
Apellidos: {apellidos}
Correo: {correo}
Edad: {edad}
Programa: {programa}

Pregunta 1 - Eleccion de carrera:
{resp1}

Pregunta 2 - Experiencia relacionada:
{resp2}

Pregunta 3 - Uso futuro del aprendizaje:
{resp3}"""
        
        missing_responses = []
        if resp1 == 'Sin respuesta': missing_responses.append('Respuesta 1')
        if resp2 == 'Sin respuesta': missing_responses.append('Respuesta 2')
        if resp3 == 'Sin respuesta': missing_responses.append('Respuesta 3')
        
        if missing_responses:
            results.append({'success': False, 'registro_numero': idx + 1, 'nombre': nombre, 'apellidos': apellidos, 'correo': correo, 'error': f"Campos faltantes: {', '.join(missing_responses)}"})
            continue
        
        analysis = analyze_admission_form(form_text)
        result = {'registro_numero': idx + 1, 'nombre': nombre, 'apellidos': apellidos, 'correo': correo, 'success': analysis.get('success', False)}
        if analysis.get('success'):
            result['analysis'] = analysis
        else:
            result['error'] = analysis.get('error', 'Error desconocido')
        results.append(result)
    
    return results


def generate_excel_report(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados Analisis SDT"
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    headers = ['N', 'Nombre', 'Apellidos', 'Correo', 'Calif. Real', 'Calif. /20',
        'R1 Punt.', 'R1 Justificacion', 'R1 Tipo', 'R2 Punt.', 'R2 Justificacion', 'R2 Tipo',
        'R3 Punt.', 'R3 Justificacion', 'R3 Tipo', 'Nivel General', 'Autonomia', 'Competencia', 'Relacion', 'Recomendaciones']
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    
    widths = [5, 15, 15, 25, 10, 10, 8, 50, 15, 8, 50, 15, 8, 50, 15, 20, 30, 30, 30, 50]
    for col_num, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = width
    
    for result in results:
        a = result.get('analysis', {})
        ws.append([
            result.get('registro_numero', ''), result.get('nombre', ''), result.get('apellidos', ''), result.get('correo', ''),
            a.get('calificacion_real', ''), a.get('calificacion_sobre_20', ''),
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


def main():
    st.markdown("""
    <div class="app-header">
        <h1 class="app-title">Sistema de Analisis de Admision</h1>
        <p class="app-subtitle">Analisis Motivacional basado en la Teoria de la Autodeterminacion (SDT)</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='background: #FFFFFF; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;'>
            <p style='margin: 0; color: #334155; font-size: 0.95rem;'>
                <strong style='color: #1E3A8A;'>Analisis Individual:</strong> PDF, DOCX, TXT |
                <strong style='color: #1E3A8A;'>Analisis Masivo:</strong> XLSX, XLS, CSV
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Seleccionar Archivo", type=['pdf', 'docx', 'doc', 'txt', 'xlsx', 'xls', 'csv'], help="Arrastra el archivo o haz clic para seleccionar", label_visibility="collapsed")
    
    if uploaded_file:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        file_size = uploaded_file.size / (1024 * 1024)
        is_batch = file_extension in ['xlsx', 'xls', 'csv']
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"<div class='info-box'><strong>{uploaded_file.name}</strong></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='info-box'>{file_size:.2f} MB</div>", unsafe_allow_html=True)
        with col3:
            mode = "Modo Masivo" if is_batch else "Modo Individual"
            badge = "badge-info" if is_batch else "badge-success"
            st.markdown(f"<div class='info-box'><span class='status-badge {badge}'>{mode}</span></div>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            analyze_button = st.button("Iniciar Analisis", type="primary", use_container_width=True)
        
        if analyze_button:
            with st.spinner("Procesando analisis..."):
                if is_batch:
                    df = read_excel_file(uploaded_file)
                    if df is not None:
                        st.success(f"{len(df)} registros detectados")
                        col_map = build_column_map(df)
                        with st.expander("Mapeo de columnas detectado", expanded=False):
                            for logical, real in col_map.items():
                                st.markdown(f"**{logical}** -> {'`' + real + '`' if real else 'No encontrada'}")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        results = process_excel_records(df, progress_bar, status_text)
                        status_text.markdown("**Analisis completado**")
                        progress_bar.progress(1.0)
                        
                        st.session_state['batch_results'] = results
                        st.markdown("<hr>", unsafe_allow_html=True)
                        st.markdown("## Resultados del Analisis Masivo")
                        
                        success_count = sum(1 for r in results if r.get('success'))
                        avg_score = sum(float(r['analysis']['calificacion_sobre_20']) for r in results if r.get('success') and r.get('analysis', {}).get('calificacion_sobre_20')) / success_count if success_count > 0 else 0
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Total", len(results))
                        c2.metric("Exitosos", success_count)
                        c3.metric("Errores", len(results) - success_count)
                        c4.metric("Promedio", f"{avg_score:.2f}/20")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        excel_buffer = generate_excel_report(results)
                        col1, col2, col3 = st.columns([1, 2, 1])
                        with col2:
                            st.download_button("Descargar Reporte Excel", data=excel_buffer, file_name=f"Analisis_SDT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("### Detalle por Postulante")
                        
                        for i, result in enumerate(results):
                            nombre_d = result.get('nombre', 'N/A')
                            apellidos_d = result.get('apellidos', 'N/A')
                            correo_d = result.get('correo', 'N/A')
                            parts = []
                            if apellidos_d != 'N/A': parts.append(apellidos_d)
                            if nombre_d != 'N/A': parts.append(nombre_d)
                            name_label = ", ".join(parts) if parts else f"Registro {result['registro_numero']}"
                            extra = f" | {correo_d}" if correo_d != 'N/A' else ""
                            
                            with st.expander(f"**{result['registro_numero']}. {name_label}**{extra}", expanded=False):
                                if result.get('success'):
                                    analysis = result['analysis']
                                    c1, c2 = st.columns([1, 3])
                                    with c1:
                                        st.markdown(f"<div style='text-align:center;padding:1rem;'><div class='score-circle'>{analysis.get('calificacion_sobre_20', 'N/A')}</div></div>", unsafe_allow_html=True)
                                    with c2:
                                        st.markdown(f"**Calificacion Real:** {analysis.get('calificacion_real', 'N/A')}/18")
                                        st.markdown(f"**Nivel Motivacional:** {analysis.get('nivel_motivacional_general', 'N/A')}")
                                        st.markdown(f"**Regla aplicada:** {analysis.get('regla_aplicada', 'N/A')}")
                                    
                                    st.markdown("---")
                                    st.markdown("#### Evaluacion Motivacional")
                                    eval_mot = analysis.get('evaluacion_motivacional', {})
                                    c1, c2, c3 = st.columns(3)
                                    with c1:
                                        if 'eleccion_carrera' in eval_mot:
                                            st.info(f"**R1: Eleccion**\n\n{eval_mot['eleccion_carrera'].get('puntaje')}/6 - {eval_mot['eleccion_carrera'].get('tipo_motivacion')}")
                                    with c2:
                                        if 'experiencia_relacionada' in eval_mot:
                                            st.info(f"**R2: Experiencia**\n\n{eval_mot['experiencia_relacionada'].get('puntaje')}/6 - {eval_mot['experiencia_relacionada'].get('tipo_motivacion')}")
                                    with c3:
                                        if 'uso_futuro' in eval_mot:
                                            st.info(f"**R3: Proyeccion**\n\n{eval_mot['uso_futuro'].get('puntaje')}/6 - {eval_mot['uso_futuro'].get('tipo_motivacion')}")
                                    
                                    st.markdown("#### Justificaciones y Reglas Aplicadas")
                                    for key, label in [('eleccion_carrera', 'R1'), ('experiencia_relacionada', 'R2'), ('uso_futuro', 'R3')]:
                                        if key in eval_mot:
                                            st.markdown(f"**{label}:** {eval_mot[key].get('justificacion', 'N/A')}")
                                    
                                    if 'necesidades_psicologicas' in analysis:
                                        st.markdown("#### Necesidades Psicologicas")
                                        nec = analysis['necesidades_psicologicas']
                                        c1, c2, c3 = st.columns(3)
                                        c1.success(f"**Autonomia**\n\n{nec.get('autonomia', 'N/A')}")
                                        c2.success(f"**Competencia**\n\n{nec.get('competencia', 'N/A')}")
                                        c3.success(f"**Relacion**\n\n{nec.get('relacion', 'N/A')}")
                                    
                                    if 'recomendaciones' in analysis:
                                        st.markdown("#### Recomendaciones")
                                        st.info(analysis['recomendaciones'])
                                else:
                                    st.error(f"**Error:** {result.get('error', 'Desconocido')}")
                
                else:
                    text_content = None
                    if file_extension == 'pdf': text_content = extract_text_from_pdf(uploaded_file)
                    elif file_extension in ['docx', 'doc']: text_content = extract_text_from_docx(uploaded_file)
                    elif file_extension == 'txt': text_content = extract_text_from_txt(uploaded_file)
                    
                    if text_content and text_content.strip():
                        st.success(f"Texto extraido ({len(text_content)} caracteres)")
                        analysis = analyze_admission_form(text_content)
                        
                        if analysis.get('success'):
                            st.markdown("<hr>", unsafe_allow_html=True)
                            st.markdown("## Resultado del Analisis Individual")
                            info = analysis.get('informacion_extraida', {})
                            c1, c2, c3 = st.columns(3)
                            c1.info(f"**Nombre**\n\n{info.get('nombre', 'N/A')}")
                            c2.info(f"**Edad**\n\n{info.get('edad', 'N/A')}")
                            c3.info(f"**Programa**\n\n{info.get('programa', 'N/A')}")
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            c1, c2 = st.columns([1, 2])
                            with c1:
                                st.markdown(f"<div style='text-align:center;padding:2rem;'><div class='score-circle'>{analysis.get('calificacion_sobre_20', 'N/A')}</div><p style='color:#64748B;margin-top:1rem;font-weight:600;'>Calificacion Final</p></div>", unsafe_allow_html=True)
                            with c2:
                                st.metric("Calificacion Real", f"{analysis.get('calificacion_real', 'N/A')}/18")
                                st.metric("Nivel Motivacional", analysis.get('nivel_motivacional_general', 'N/A'))
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.markdown("### Evaluacion Motivacional Detallada")
                            eval_mot = analysis.get('evaluacion_motivacional', {})
                            for key, label in [('eleccion_carrera', 'Eleccion de Carrera'), ('experiencia_relacionada', 'Experiencia Relacionada'), ('uso_futuro', 'Proyeccion Futura')]:
                                if key in eval_mot:
                                    item = eval_mot[key]
                                    with st.expander(f"**{label}** - Puntaje: {item.get('puntaje')}/6 - {item.get('tipo_motivacion')}", expanded=True):
                                        st.markdown(f"**Justificacion:** {item.get('justificacion')}")
                            
                            if 'necesidades_psicologicas' in analysis:
                                st.markdown("### Necesidades Psicologicas (SDT)")
                                nec = analysis['necesidades_psicologicas']
                                c1, c2, c3 = st.columns(3)
                                c1.success(f"**Autonomia**\n\n{nec.get('autonomia', 'N/A')}")
                                c2.success(f"**Competencia**\n\n{nec.get('competencia', 'N/A')}")
                                c3.success(f"**Relacion**\n\n{nec.get('relacion', 'N/A')}")
                            
                            if 'recomendaciones' in analysis:
                                st.markdown("### Recomendaciones")
                                st.info(analysis['recomendaciones'])
                            
                            st.markdown("<hr>", unsafe_allow_html=True)
                            st.caption(f"**Archivo:** {uploaded_file.name} | **Tokens:** {analysis.get('tokens_used', 'N/A')} | **Procesado:** {datetime.fromisoformat(analysis.get('timestamp')).strftime('%d/%m/%Y %H:%M:%S')}")
                        else:
                            st.error(f"**Error:** {analysis.get('error', 'Desconocido')}")
                    else:
                        st.error("No se pudo extraer texto del archivo")
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<div style='text-align:center;padding:2rem;color:#94A3B8;font-size:0.85rem;'><p>Direccion de Gestion de la Informacion - Universidad Continental</p></div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()

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

# Estilos CSS profesionales y modernos
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
    
    /* Métricas mejoradas */
    div[data-testid="metric-container"] {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        border-left: 4px solid #3B82F6;
    }
    
    div[data-testid="metric-container"] label {
        color: #64748B;
        font-size: 0.875rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #1E3A8A;
        font-size: 2rem;
        font-weight: 700;
    }
    
    /* Cards de información */
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
        opacity: 0.8;
    }
    
    /* Progress bar personalizada */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
    }
    
    /* Expander mejorado */
    .streamlit-expanderHeader {
        background: white;
        border-radius: 8px;
        font-weight: 600;
        color: #1E3A8A;
        border: 1px solid #E2E8F0;
    }
    
    .streamlit-expanderHeader:hover {
        background: #F8FAFC;
        border-color: #3B82F6;
    }
    
    /* Download button especial */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: white;
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
    
    /* Alerts personalizados */
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid;
    }
    
    /* Info boxes */
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
    
    /* Separator line */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, #3B82F6, transparent);
        margin: 2rem 0;
    }
    
    /* Caption mejorado */
    .stCaption {
        color: #64748B !important;
        font-size: 0.875rem !important;
    }
    
    /* Spinner personalizado */
    .stSpinner > div {
        border-top-color: #3B82F6 !important;
    }
    
    /* Tablas */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Container principal */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Success/Warning/Error boxes específicos */
    .success-box {
        background: #D1FAE5;
        border-left: 4px solid #059669;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .warning-box {
        background: #FEF3C7;
        border-left: 4px solid #F59E0B;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .info-box {
        background: #DBEAFE;
        border-left: 4px solid #3B82F6;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
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

═══════════════════════════════════════════════════════════════════════
RÚBRICA DETALLADA POR NIVEL
═══════════════════════════════════════════════════════════════════════

NIVEL 6 - MOTIVACIÓN INTRÍNSECA:
Criterio general: Interés genuino, disfrute, curiosidad o satisfacción personal.
- Elección de carrera: Se basa en el agrado inherente por los contenidos o actividades propias de la carrera
- Experiencia: Disfrute del proceso, interés espontáneo, sensación de flujo. El valor está en realizarla, no en resultados
- Uso futuro: Desea aplicar lo aprendido por interés y disfrute personal, motivación espontánea
Indicadores lingüísticos: "me gusta", "lo disfruto", "me interesa mucho", "me resulta entretenido", "me apasiona"
NO asignar si: el interés se justifica por utilidad, resultados, metas, identidad o impacto social

NIVEL 5 - REGULACIÓN INTEGRADA:
Criterio general: Coherencia con identidad, valores centrales y proyecto de vida.
- Elección de carrera: La carrera es coherente con quién es y quién desea ser
- Experiencia: Fortalece identidad y sentido de coherencia personal, se integra a proyecto de vida
- Uso futuro: Integra el conocimiento a su identidad futura, aplicarlo es parte de quién quiere ser
Indicadores lingüísticos: "es coherente con mis valores", "encaja con mi forma de desarrollarme", "es parte de mi proyecto"
NO asignar si: solo hay disfrute (→6), solo utilidad (→4), o se infiere sin declaración explícita

NIVEL 4 - REGULACIÓN IDENTIFICADA:
Criterio general: Reconoce importancia y utilidad personal para metas significativas.
- Elección de carrera: Importante para desarrollo personal, académico o profesional, aunque no siempre la disfrute
- Experiencia: Valorada porque permitió aprender, desarrollar habilidades o avanzar hacia metas relevantes
- Uso futuro: Útil para metas personales importantes (trabajo, desarrollo, contribución social)
Indicadores lingüísticos: "es importante para mí", "me permite desarrollarme", "me ayuda a lograr mis metas"
NO asignar si: solo hay recompensas externas (→2), se menciona identidad/proyecto vital (→5)

NIVEL 3 - REGULACIÓN INTROYECTADA:
Criterio general: Presión interna, necesidad de validación, evitar emociones negativas.
- Elección de carrera: Justifica por culpa, orgullo, evitar decepcionar, necesidad de validación
- Experiencia: Significativa por orgullo, autoexigencia, evitar culpa, demostrarse capacidad
- Uso futuro: Aplicar para no fallar, cumplir expectativas internas, evitar sentirse insuficiente
Indicadores lingüísticos: "sentía que debía", "no quería fallar", "quería demostrar", "me sentiría mal si no"
NO asignar si: hay demandas externas explícitas (→2), hay valor personal o metas (→4)

NIVEL 2 - REGULACIÓN EXTERNA:
Criterio general: Recompensas externas, demandas sociales, control externo.
- Elección de carrera: Por dinero, prestigio, seguridad laboral, expectativas familiares o sociales
- Experiencia: Importante por notas, premios, reconocimiento, beneficios materiales
- Uso futuro: Orientado a empleo, dinero, estabilidad, cumplir requisitos externos
Indicadores lingüísticos: "tiene buena salida laboral", "da estabilidad", "mis padres querían", "para conseguir trabajo"
NO asignar si: hay culpa/orgullo (→3), valor personal explícito (→4)

NIVEL 1 - AMOTIVACIÓN:
Criterio general: Incapaz de dar razón clara, desinterés, falta de control.
- Elección de carrera: Postula por inercia, azar, resignación, no sabe por qué, percibe decisión fuera de control
- Experiencia: Apatía, desgano, aburrimiento, incompetencia, actividad pasiva o mecánica
- Uso futuro: No visualiza futuro profesional, duda de ejercer, considera aprendizaje inútil
Indicadores lingüísticos: "no lo tengo claro", "no sé por qué", "me da igual", "me obligaron", "solo porque toca"
SE ASIGNA si ocurre AL MENOS UNA condición de amotivación

═══════════════════════════════════════════════════════════════════════
REGLAS DE ASIGNACIÓN (OBLIGATORIAS)
═══════════════════════════════════════════════════════════════════════

1. EVIDENCIA EXPLÍCITA: Solo asignar niveles con indicadores EXPLÍCITOS en el texto. NO hacer inferencias implícitas.
2. UNA SOLA CLASIFICACIÓN: Cada respuesta se clasifica en UN solo nivel por criterio.
3. NIVEL MENOS AUTÓNOMO: Si coexisten indicadores de varios niveles → asignar el nivel INFERIOR (menos autónomo).
4. CONDICIONES MÍNIMAS: Para asignar un nivel, TODAS sus condiciones mínimas deben cumplirse. Si falta una → evaluar nivel inferior.
5. EVALUACIÓN INDEPENDIENTE: Cada pregunta (P1, P2, P3) se evalúa de manera INDEPENDIENTE.
6. OBJETO PERTINENTE: Intereses generales no vinculados al objeto de la pregunta NO justifican motivación autónoma.
7. PROCESO vs RESULTADO: Motivación intrínseca = foco en PROCESO; Regulación identificada = foco en RESULTADOS
8. INTERÉS PROSOCIAL: Si el interés está dirigido a impactos, utilidad social o resultados → NO es intrínseco (→4 o 5)
9. NOTAS ACLARATORIAS:
   P2: Asignar intrínseca si hay disfrute genuino del proceso, aun cuando confirme vocación posteriormente
   P3: Vinculación explícita existe cuando la acción corresponde directamente al ejercicio profesional
10. PERFIL FINAL:
    REGLA BASE: Perfil = min(P1, P2, P3)
    EXCEPCIÓN 2-de-3: Si dos coinciden y tercera está 1 nivel abajo → perfil = nivel coincidente
    Ejemplos: (4,4,3)→4, (5,5,4)→5, (6,6,5)→6
    No válidos: (6,5,4)→4, (5,5,3)→3
    SEGURIDAD: Si alguna = 1 → perfil final máximo = 2

═══════════════════════════════════════════════════════════════════════
CONDICIONES MÍNIMAS POR PREGUNTA
═══════════════════════════════════════════════════════════════════════

P1: ELECCIÓN DE CARRERA
[6] ✓ Interés/disfrute explícito ✓ Objeto: carrera ✓ Sin utilidad/metas/empleo
[5] ✓ Principios/valores/vocación ✓ Coherencia proyecto vida ✓ Principios estables
[4] ✓ Valoración explícita ✓ Objeto: carrera ✓ Metas significativas
[3] ✓ Presión interna ✓ Autoevaluación ✓ Sin recompensas externas
[2] ✓ Recompensa/demanda externa ✓ Objeto: carrera ✓ Instrumental
[1] AL MENOS UNA: ✓ Sin razón ✓ Desinterés ✓ Fuera de control

P2: EXPERIENCIA RELACIONADA
[6] ✓ Disfrute/interés explícito ✓ Objeto: actividad ✓ Foco en PROCESO
[5] ✓ Conectada con identidad ✓ Refuerzo propósito ✓ Integración proyecto
[4] ✓ Valoración explícita ✓ Aprendizaje/habilidades ✓ Metas relevantes
[3] ✓ Orgullo/autoexigencia/culpa ✓ Demostrarse capacidad ✓ Sin recompensas externas
[2] ✓ Recompensa externa ✓ Valorada por resultado ✓ Control externo
[1] AL MENOS UNA: ✓ Apatía ✓ Sin sentido ✓ Pasiva

P3: USO FUTURO
[6] ✓ Interés por aplicar/aprender ✓ Objeto: conocimiento ✓ Espontánea
[5] ✓ Vinculado identidad futura ✓ Coherencia proyecto ✓ Parte de "quién ser"
[4] ✓ Utilidad explícita ✓ Metas relevantes ✓ Elección voluntaria
[3] ✓ Para no fallar ✓ Presión interna ✓ Evaluación del yo
[2] ✓ Recompensas externas ✓ Empleo/dinero ✓ Control externo
[1] AL MENOS UNA: ✓ No visualiza ✓ Duda ejercer ✓ Inútil

═══════════════════════════════════════════════════════════════════════
ESTRUCTURA JSON
═══════════════════════════════════════════════════════════════════════

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
- perfil = aplicar reglas min/2-de-3/seguridad

IMPORTANTE:
✓ Solo JSON ✓ Sin markdown ✓ Comillas dobles ✓ Evidencia explícita ✓ Solo motivación
"""

    user_prompt = f"""PREGUNTAS DEL FORMULARIO:
1. ¿Qué características de esta carrera llamaron tu atención y cuál es la razón principal por la que decidiste postular a ella?
2. Relata una experiencia donde hayas puesto en práctica habilidades relacionadas con esta carrera. Describe cómo te sentiste mientras realizabas dicha actividad y qué descubriste de tu vocación profesional.
3. Imagina que ya terminaste tus estudios. ¿Cómo aplicarías lo aprendido en tu formación profesional y qué impactos te gustaría lograr?

═══════════════════════════════════════════════════════════════════════
FORMULARIO:
═══════════════════════════════════════════════════════════════════════

{text_content}

═══════════════════════════════════════════════════════════════════════
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

# Función para procesar registros de Excel
def process_excel_records(df, progress_bar, status_text):
    results = []
    total = len(df)
    
    required_fields = ['Respuesta 1', 'Respuesta 2', 'Respuesta 3']
    
    for idx, row in df.iterrows():
        status_text.markdown(f"**Procesando:** Registro {idx + 1} de {total}")
        progress_bar.progress((idx + 1) / total)
        
        form_text = f"""
Nombre: {row.get('Nombre', 'N/A')}
Apellidos: {row.get('Apellidos', 'N/A')}
Correo: {row.get('Correo electrónico', 'N/A')}
Edad: {row.get('Edad', 'N/A')}
Programa: {row.get('Programa', 'N/A')}

Pregunta 1: ¿Por qué elegiste esta carrera?
Respuesta 1: {row.get('Respuesta 1', 'Sin respuesta')}

Pregunta 2: ¿Qué experiencia tienes relacionada con esta carrera?
Respuesta 2: {row.get('Respuesta 2', 'Sin respuesta')}

Pregunta 3: ¿Cómo planeas usar lo que aprendas?
Respuesta 3: {row.get('Respuesta 3', 'Sin respuesta')}
"""
        
        missing_fields = [field for field in required_fields if pd.isna(row.get(field)) or str(row.get(field)).strip() == '']
        
        if missing_fields:
            results.append({
                'success': False,
                'registro_numero': idx + 1,
                'nombre': row.get('Nombre', 'N/A'),
                'apellidos': row.get('Apellidos', 'N/A'),
                'correo': row.get('Correo electrónico', 'N/A'),
                'error': f"Campos faltantes: {', '.join(missing_fields)}"
            })
            continue
        
        analysis = analyze_admission_form(form_text)
        
        result = {
            'registro_numero': idx + 1,
            'nombre': row.get('Nombre', 'N/A'),
            'apellidos': row.get('Apellidos', 'N/A'),
            'correo': row.get('Correo electrónico', 'N/A'),
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

# INTERFAZ PRINCIPAL
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
        <div style='background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;'>
            <p style='margin: 0; color: #64748B; font-size: 0.95rem;'>
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
                    # PROCESAMIENTO MASIVO
                    df = read_excel_file(uploaded_file)
                    
                    if df is not None:
                        st.success(f"✅ {len(df)} registros detectados")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        results = process_excel_records(df, progress_bar, status_text)
                        
                        status_text.markdown("✅ **Análisis completado exitosamente**")
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
        <p>Sistema de Análisis de Admisión • Universidad Continental</p>
        <p>Basado en la Teoría de la Autodeterminación (SDT) de Ryan y Deci</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

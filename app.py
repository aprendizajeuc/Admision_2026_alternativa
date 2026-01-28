import streamlit as st
import os
from openai import OpenAI

st.title("🔍 Diagnóstico de API Key")

st.markdown("---")
st.header("1️⃣ Verificar qué API Key se está usando")

# Intentar obtener la API Key
api_key = None
source = "No encontrada"

# Método 1: Streamlit Secrets
try:
    api_key = st.secrets["OPENAI_API_KEY"]
    source = "Streamlit Secrets (st.secrets)"
    st.success(f"✅ API Key encontrada en: **{source}**")
except Exception as e:
    st.warning(f"⚠️ No se encontró en Streamlit Secrets: {str(e)}")

# Método 2: Variable de entorno
if not api_key:
    try:
        api_key = os.getenv('OPENAI_API_KEY')
        if api_key:
            source = "Variable de entorno (os.getenv)"
            st.success(f"✅ API Key encontrada en: **{source}**")
        else:
            st.error("❌ No se encontró en variables de entorno")
    except Exception as e:
        st.error(f"❌ Error al buscar en variables de entorno: {str(e)}")

st.markdown("---")
st.header("2️⃣ Información de la API Key")

if api_key:
    # Mostrar información de la key
    st.info(f"""
    **Fuente:** {source}
    
    **Primeros 10 caracteres:** `{api_key[:10]}...`
    
    **Últimos 4 caracteres:** `...{api_key[-4:]}`
    
    **Longitud total:** {len(api_key)} caracteres
    
    **Formato esperado:** sk-proj-... (muy largo, ~164-200 caracteres)
    """)
    
    # Verificar formato
    if api_key.startswith('sk-proj-'):
        st.success("✅ Formato correcto: Empieza con 'sk-proj-'")
    else:
        st.error(f"❌ Formato incorrecto: Empieza con '{api_key[:10]}...'")
    
    if len(api_key) < 100:
        st.error(f"❌ Key muy corta ({len(api_key)} caracteres). Debería tener ~164-200 caracteres")
    else:
        st.success(f"✅ Longitud correcta ({len(api_key)} caracteres)")
    
else:
    st.error("❌ No se pudo obtener ninguna API Key")
    st.stop()

st.markdown("---")
st.header("3️⃣ Probar conexión con OpenAI")

if st.button("🧪 Probar API Key", type="primary"):
    with st.spinner("Probando conexión con OpenAI..."):
        try:
            client = OpenAI(api_key=api_key)
            
            # Hacer una petición simple
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "user", "content": "Di solo la palabra 'Hola'"}
                ],
                max_tokens=10
            )
            
            result = response.choices[0].message.content
            
            st.success("✅ ¡API Key funciona correctamente!")
            st.info(f"**Respuesta de OpenAI:** {result}")
            st.balloons()
            
        except Exception as e:
            st.error(f"❌ Error al probar la API Key:")
            st.code(str(e))
            
            # Analizar el error
            error_str = str(e)
            
            if "401" in error_str:
                st.error("""
                **Error 401: API Key Inválida**
                
                Posibles causas:
                1. La API Key está mal copiada (caracteres faltantes)
                2. La API Key fue revocada en OpenAI
                3. Hay espacios o caracteres extra
                """)
                
                # Mostrar qué parte de la key está fallando
                if "Incorrect API key provided:" in error_str:
                    st.warning("La key que está llegando a OpenAI está incorrecta")
                    
            elif "429" in error_str:
                st.warning("Límite de rate excedido. La key funciona pero has hecho muchas peticiones.")
                
            elif "quota" in error_str.lower():
                st.warning("Sin créditos en OpenAI. Agrega método de pago en platform.openai.com/account/billing")

st.markdown("---")
st.header("4️⃣ Debug: Ver contenido de Secrets")

if st.checkbox("🔓 Mostrar todos los Secrets disponibles (CUIDADO: No hacer esto en producción)"):
    st.warning("⚠️ Esto mostrará información sensible. Solo para debug.")
    try:
        st.json(dict(st.secrets))
    except:
        st.error("No se pudieron obtener los secrets")

st.markdown("---")
st.header("5️⃣ Instrucciones de Solución")

st.info("""
### Si la API Key NO funciona:

1. **Ve a OpenAI:**
   - URL: https://platform.openai.com/api-keys
   
2. **Crea una NUEVA API Key:**
   - Click en "+ Create new secret key"
   - Nombre: "Streamlit Test"
   - Click "Create"
   - **COPIA LA KEY COMPLETA** (solo se muestra una vez)
   
3. **Actualiza en Streamlit Cloud:**
   - Ve a tu app en share.streamlit.io
   - Settings → Secrets
   - BORRA TODO el contenido
   - Pega exactamente:
   ```
   OPENAI_API_KEY = "sk-proj-TU_NUEVA_KEY_AQUI"
   ```
   - Click "Save"
   - Click "Reboot app"
   
4. **Recarga esta página de diagnóstico**
   - Debería mostrar la nueva key
   - Prueba la conexión con el botón de arriba
""")

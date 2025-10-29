"""
EqualityMomentum - Aplicación Web
Procesamiento de Registros Retributivos

Versión web sin almacenamiento de datos
"""

import streamlit as st
import io
import sys
from pathlib import Path
from datetime import datetime

# Agregar la ruta de 04_SCRIPTS al path para importar el generador optimizado
scripts_path = Path(__file__).parent / '04_SCRIPTS'
if str(scripts_path) not in sys.path:
    sys.path.insert(0, str(scripts_path))

# Importar módulos
from core.procesador import ProcesadorRegistroRetributivo
from generar_informe_optimizado import GeneradorInformeOptimizado


# Configuración de la página
st.set_page_config(
    page_title="EqualityMomentum",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado con estilo corporativo de EqualityMomentum
st.markdown("""
    <style>
    /* Importar fuentes de Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Lusitana:wght@400;700&family=Work+Sans:wght@300;400;500;600;700&display=swap');
    
    /* Variables de colores corporativos */
    :root {
        --azul-corporativo: #1f3c89;
        --naranja-corporativo: #ff5c39;
        --blanco-corporativo: #ffffff;
    }
    
    /* Estilos generales */
    .main {
        padding: 2rem;
        font-family: 'Work Sans', sans-serif;
        font-size: 18px;
    }
    
    /* Títulos con tipografía Lusitana */
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Lusitana', serif !important;
        color: var(--azul-corporativo) !important;
    }
    
    h1 {
        font-size: 48px !important;
    }
    
    h2 {
        font-size: 36px !important;
    }
    
    h3 {
        font-size: 28px !important;
    }
    
    /* Textos con Work Sans */
    p, div, span, label {
        font-family: 'Work Sans', sans-serif;
        font-size: 18px;
    }
    
    /* Botones con colores corporativos */
    .stButton>button {
        width: 100%;
        background-color: var(--azul-corporativo);
        color: var(--blanco-corporativo);
        font-family: 'Work Sans', sans-serif;
        font-weight: 600;
        font-size: 18px;
        padding: 0.75rem;
        border-radius: 8px;
        border: none;
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        background-color: var(--naranja-corporativo);
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(31, 60, 137, 0.3);
    }
    
    /* Botón primario */
    .stButton>button[kind="primary"] {
        background-color: var(--naranja-corporativo);
    }
    
    .stButton>button[kind="primary"]:hover {
        background-color: var(--azul-corporativo);
    }
    
    /* Sección de carga con estilo corporativo */
    .upload-section {
        border: 2px dashed var(--azul-corporativo);
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: rgba(31, 60, 137, 0.05);
    }
    
    /* Cajas informativas */
    .info-box {
        background-color: rgba(31, 60, 137, 0.1);
        border-left: 4px solid var(--azul-corporativo);
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
        font-family: 'Work Sans', sans-serif;
    }
    
    .success-box {
        background-color: rgba(76, 175, 80, 0.1);
        border-left: 4px solid #4caf50;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
        font-family: 'Work Sans', sans-serif;
    }
    
    .warning-box {
        background-color: rgba(255, 92, 57, 0.1);
        border-left: 4px solid var(--naranja-corporativo);
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
        font-family: 'Work Sans', sans-serif;
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: rgba(31, 60, 137, 0.05);
    }
    
    /* Enlaces */
    a {
        color: var(--naranja-corporativo);
        text-decoration: none;
    }
    
    a:hover {
        color: var(--azul-corporativo);
        text-decoration: underline;
    }
    
    /* Separadores */
    hr {
        border-color: var(--azul-corporativo);
        opacity: 0.3;
    }
    </style>
""", unsafe_allow_html=True)


def limpiar_sesion():
    """Limpia los datos de la sesión para liberar memoria"""
    keys_to_clear = ['archivo_procesado', 'informe_word', 'nombre_archivo', 'estadisticas']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]


def main():
    # Header con logotipo corporativo
    st.markdown("""
        <div style="display: flex; align-items: center; justify-content: center; margin-bottom: 2rem;">
            <img src="https://equalitymomentum.com/wp-content/uploads/2024/04/equality-momentum-imagotipo.svg" 
                 alt="EqualityMomentum" 
                 style="max-width: 400px; width: 100%; height: auto;">
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div style="text-align: center; margin-bottom: 2rem;">
            <h1 style="font-family: 'Lusitana', serif; font-size: 48px; color: #1f3c89; margin-bottom: 0.5rem;">
                Procesamiento de Registros Retributivos
            </h1>
            <p style="font-family: 'Work Sans', sans-serif; font-size: 18px; color: #666;">
                Herramienta profesional para análisis de igualdad retributiva
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Banner de privacidad
    st.markdown("""
        <div class="info-box">
            <strong>🔒 Privacidad Garantizada:</strong> Todos los archivos se procesan en memoria.
            No se almacenan datos en el servidor. Los archivos se eliminan automáticamente al cerrar la sesión.
        </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # Sidebar con información
    with st.sidebar:
        st.header("ℹ️ Información")

        st.markdown("""
        ### Cómo usar:
        1. Selecciona el tipo de archivo
        2. Sube tu archivo Excel
        3. Procesa los datos
        4. Descarga los resultados

        ### Formatos admitidos:
        - **General**: Formato estándar con hoja "BASE GENERAL"
        - **Triodos**: Formato bancario (protegido con contraseña)

        ### Archivos generados:
        - **Excel**: Datos procesados con columnas equiparadas
        - **Word**: Informe completo con gráficos y tablas
        """)

        st.markdown("---")

        st.markdown("""
        ### 🛡️ Seguridad:
        - Sin base de datos
        - Sin logs con datos personales
        - Procesamiento en memoria RAM
        - Limpieza automática
        """)

        if st.button("🗑️ Limpiar Sesión"):
            limpiar_sesion()
            st.success("Sesión limpiada correctamente")
            st.rerun()

    # Contenido principal
    st.header("1️⃣ Configuración")

    # Selector de tipo de archivo y opciones
    col1, col2 = st.columns(2)
    with col1:
        tipo_archivo = st.selectbox(
            "Tipo de archivo:",
            options=["General", "Triodos"],
            help="Selecciona el formato de tu archivo Excel"
        )

    with col2:
        accion = st.selectbox(
            "Acción a realizar:",
            options=["Ambas", "Procesar Datos", "Generar Informe"],
            help="Selecciona qué operación deseas realizar"
        )

    # Opciones adicionales
    col1, col2 = st.columns(2)
    with col1:
        archivo_protegido = st.checkbox(
            "¿El archivo tiene contraseña?",
            value=(tipo_archivo == "Triodos"),
            help="Marca esta casilla si el archivo Excel está protegido"
        )

    with col2:
        if archivo_protegido:
            password = st.text_input(
                "Contraseña del archivo:",
                value="Triodos2025" if tipo_archivo == "Triodos" else "",
                type="password",
                help="Contraseña para desbloquear el archivo Excel"
            )
        else:
            password = None

    st.markdown("---")

    # Sección de carga de archivo
    st.header("2️⃣ Cargar Archivo")

    st.markdown('<div class="upload-section">', unsafe_allow_html=True)

    archivo_subido = st.file_uploader(
        "Arrastra tu archivo Excel aquí o haz clic para seleccionar",
        type=['xlsx', 'xls'],
        help="Tamaño máximo: 50MB",
        label_visibility="collapsed"
    )

    st.markdown('</div>', unsafe_allow_html=True)

    if archivo_subido is not None:
        st.success(f"✅ Archivo cargado: **{archivo_subido.name}** ({archivo_subido.size / 1024:.2f} KB)")

        # Validar tamaño (50MB máximo)
        if archivo_subido.size > 50 * 1024 * 1024:
            st.error("❌ El archivo es demasiado grande. Tamaño máximo: 50MB")
            return

        st.markdown("---")

        # Botón de procesamiento
        st.header("3️⃣ Procesar Datos")

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            # Cambiar el texto del botón según la acción
            texto_boton = {
                "Ambas": "🚀 Procesar y Generar Informe",
                "Procesar Datos": "📊 Procesar Datos",
                "Generar Informe": "📄 Generar Informe"
            }

            if st.button(texto_boton[accion], type="primary"):
                with st.spinner(f"{accion}... Esto puede tardar unos segundos."):
                    try:
                        # Leer archivo como bytes
                        archivo_bytes = archivo_subido.read()

                        excel_procesado = None
                        informe_word = None

                        # PASO 1: Procesar datos (si corresponde)
                        if accion in ["Ambas", "Procesar Datos"]:
                            with st.spinner("📊 Procesando datos..."):
                                procesador = ProcesadorRegistroRetributivo()

                                # Procesar según tipo y protección
                                if tipo_archivo == "Triodos" or archivo_protegido:
                                    excel_procesado = procesador.procesar_excel_triodos(
                                        archivo_bytes,
                                        password=password if password else "Triodos2025"
                                    )
                                else:
                                    excel_procesado = procesador.procesar_excel_general(archivo_bytes)

                                st.success("✅ Datos procesados correctamente")

                        # PASO 2: Generar informe (si corresponde)
                        if accion in ["Ambas", "Generar Informe"]:
                            with st.spinner("📄 Generando informe Word..."):
                                # Si no se procesaron los datos, usar el archivo original
                                if excel_procesado is None:
                                    excel_para_informe = io.BytesIO(archivo_bytes)
                                else:
                                    excel_procesado.seek(0)
                                    excel_para_informe = excel_procesado

                                # Usar el generador optimizado
                                generador = GeneradorInformeOptimizado()
                                
                                # Cargar datos desde BytesIO
                                if generador.cargar_datos_desde_bytes(excel_para_informe):
                                    # Generar informe (por defecto CONSOLIDADO)
                                    informe_word = generador.generar_informe_bytes('CONSOLIDADO')
                                    
                                    if informe_word:
                                        st.success("✅ Informe generado correctamente")
                                    else:
                                        st.error("❌ Error al generar el informe Word")
                                        informe_word = None
                                else:
                                    st.error("❌ Error al cargar los datos para el informe")
                                    informe_word = None

                        # Guardar en session_state
                        if excel_procesado:
                            excel_procesado.seek(0)
                            st.session_state['archivo_procesado'] = excel_procesado

                        if informe_word:
                            st.session_state['informe_word'] = informe_word

                        st.session_state['nombre_archivo'] = archivo_subido.name.replace('.xlsx', '').replace('.xls', '')
                        st.session_state['accion_realizada'] = accion

                        # Calcular estadísticas básicas (si se procesó)
                        if excel_procesado:
                            import pandas as pd
                            excel_procesado.seek(0)
                            df = pd.read_excel(excel_procesado, sheet_name='DATOS_PROCESADOS')
                            st.session_state['estadisticas'] = {
                                'total_registros': len(df),
                                'columnas': len(df.columns)
                            }

                        st.success(f"✅ {accion} completado exitosamente!")

                    except Exception as e:
                        st.error(f"❌ Error durante el procesamiento: {str(e)}")
                        st.exception(e)

        # Mostrar resultados si existen
        if 'archivo_procesado' in st.session_state or 'informe_word' in st.session_state:
            st.markdown("---")
            st.header("4️⃣ Resultados")

            # Estadísticas
            if 'estadisticas' in st.session_state:
                stats = st.session_state['estadisticas']

                st.markdown("""
                    <div class="success-box">
                        <strong>📊 Estadísticas del procesamiento:</strong><br>
                        • Registros procesados: <strong>{}</strong><br>
                        • Columnas generadas: <strong>{}</strong>
                    </div>
                """.format(stats['total_registros'], stats['columnas']), unsafe_allow_html=True)

            # Botones de descarga
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Determinar cuántas columnas necesitamos
            tiene_excel = 'archivo_procesado' in st.session_state
            tiene_word = 'informe_word' in st.session_state

            if tiene_excel and tiene_word:
                col1, col2 = st.columns(2)
            elif tiene_excel or tiene_word:
                col1 = st.container()
                col2 = None

            # Botón de descarga Excel
            if tiene_excel:
                with col1:
                    st.subheader("📊 Excel Procesado")
                    st.markdown("Datos con columnas equiparadas")

                    nombre_excel = f"REPORTE_{st.session_state['nombre_archivo']}_{timestamp}.xlsx"

                    st.session_state['archivo_procesado'].seek(0)
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=st.session_state['archivo_procesado'],
                        file_name=nombre_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            # Botón de descarga Word
            if tiene_word:
                with (col2 if col2 else col1):
                    st.subheader("📄 Informe Word")
                    st.markdown("Informe completo con gráficos")

                    nombre_word = f"INFORME_{st.session_state['nombre_archivo']}_{timestamp}.docx"

                    st.session_state['informe_word'].seek(0)
                    st.download_button(
                        label="⬇️ Descargar Informe",
                        data=st.session_state['informe_word'],
                        file_name=nombre_word,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

            st.markdown("""
                <div class="warning-box">
                    <strong>⚠️ Importante:</strong> Los archivos se eliminan automáticamente al cerrar esta ventana
                    o actualizar la página. Descarga tus archivos antes de salir.
                </div>
            """, unsafe_allow_html=True)

    # Footer con estilo corporativo
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; padding: 2rem 0;">
            <p style="font-family: 'Work Sans', sans-serif; color: #1f3c89; font-size: 16px; font-weight: 600; margin-bottom: 0.5rem;">
                <strong>EqualityMomentum</strong> v2.0 | Registro Retributivo
            </p>
            <p style="font-family: 'Work Sans', sans-serif; color: #666; font-size: 14px;">
                🔒 Sin almacenamiento de datos | Procesamiento en memoria
            </p>
            <p style="font-family: 'Work Sans', sans-serif; color: #999; font-size: 12px; margin-top: 1rem;">
                © 2025 EqualityMomentum - Todos los derechos reservados
            </p>
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

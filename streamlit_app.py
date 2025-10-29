"""
EqualityMomentum - Aplicación Web
Procesamiento de Registros Retributivos

Versión web sin almacenamiento de datos
"""

import streamlit as st
import io
from datetime import datetime

# Importar módulos core
from core.procesador import ProcesadorRegistroRetributivo
from core.generador import GeneradorInformes


# Configuración de la página
st.set_page_config(
    page_title="EqualityMomentum",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para mejorar la estética
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #1e4389;
        color: white;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 8px;
        border: none;
    }
    .stButton>button:hover {
        background-color: #ea5d41;
    }
    .upload-section {
        border: 2px dashed #1e4389;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
    }
    .info-box {
        background-color: #e3f2fd;
        border-left: 4px solid #1e4389;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
    }
    .success-box {
        background-color: #e8f5e9;
        border-left: 4px solid #4caf50;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
    }
    .warning-box {
        background-color: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
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
    # Header
    col1, col2 = st.columns([1, 4])
    with col1:
        st.markdown("# ⚖️")
    with col2:
        st.title("EqualityMomentum")
        st.markdown("**Procesamiento de Registros Retributivos**")

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

    # Selector de tipo de archivo
    col1, col2 = st.columns(2)
    with col1:
        tipo_archivo = st.selectbox(
            "Tipo de archivo:",
            options=["General", "Triodos"],
            help="Selecciona el formato de tu archivo Excel"
        )

    with col2:
        if tipo_archivo == "Triodos":
            password = st.text_input(
                "Contraseña del archivo:",
                value="Triodos2025",
                type="password",
                help="Contraseña para desbloquear el archivo Excel de Triodos"
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
            if st.button("🚀 Procesar Archivo", type="primary"):
                with st.spinner("Procesando datos... Esto puede tardar unos segundos."):
                    try:
                        # Leer archivo como bytes
                        archivo_bytes = archivo_subido.read()

                        # Crear procesador
                        procesador = ProcesadorRegistroRetributivo()

                        # Procesar según tipo
                        if tipo_archivo == "Triodos":
                            excel_procesado = procesador.procesar_excel_triodos(archivo_bytes, password=password)
                        else:
                            excel_procesado = procesador.procesar_excel_general(archivo_bytes)

                        # Generar informe Word
                        generador = GeneradorInformes()
                        excel_procesado.seek(0)  # Resetear puntero
                        informe_word = generador.generar_informe_completo(excel_procesado)

                        # Guardar en session_state
                        excel_procesado.seek(0)  # Resetear otra vez
                        st.session_state['archivo_procesado'] = excel_procesado
                        st.session_state['informe_word'] = informe_word
                        st.session_state['nombre_archivo'] = archivo_subido.name.replace('.xlsx', '').replace('.xls', '')

                        # Calcular estadísticas básicas
                        import pandas as pd
                        excel_procesado.seek(0)
                        df = pd.read_excel(excel_procesado, sheet_name='DATOS_PROCESADOS')
                        st.session_state['estadisticas'] = {
                            'total_registros': len(df),
                            'columnas': len(df.columns)
                        }

                        st.success("✅ Procesamiento completado exitosamente!")

                    except Exception as e:
                        st.error(f"❌ Error durante el procesamiento: {str(e)}")
                        st.exception(e)

        # Mostrar resultados si existen
        if 'archivo_procesado' in st.session_state:
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
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📊 Excel Procesado")
                st.markdown("Datos con columnas equiparadas")

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_excel = f"REPORTE_{st.session_state['nombre_archivo']}_{timestamp}.xlsx"

                st.session_state['archivo_procesado'].seek(0)
                st.download_button(
                    label="⬇️ Descargar Excel",
                    data=st.session_state['archivo_procesado'],
                    file_name=nombre_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col2:
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

    # Footer
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; color: #666; font-size: 0.9em; padding: 2rem 0;">
            <strong>EqualityMomentum</strong> v2.0 | Registro Retributivo |
            🔒 Sin almacenamiento de datos | Procesamiento en memoria
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

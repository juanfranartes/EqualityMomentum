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
import pandas as pd
import tempfile
import os

# Configuración de la página (DEBE SER LO PRIMERO)
st.set_page_config(
    page_title="EqualityMomentum",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Agregar la ruta de 04_SCRIPTS al path para importar los módulos
scripts_path = Path(__file__).parent / '04_SCRIPTS'
if str(scripts_path) not in sys.path:
    sys.path.insert(0, str(scripts_path))

# IMPORTANTE: Recargar módulos para asegurar que se usan las versiones más recientes
# Limpiar caché de módulos si existen
modulos_a_recargar = ['procesar_datos', 'procesar_datos_triodos', 'generar_informe_optimizado', 'validador_mapeo']
for modulo_nombre in modulos_a_recargar:
    if modulo_nombre in sys.modules:
        del sys.modules[modulo_nombre]

# Ahora importar las versiones frescas
try:
    from procesar_datos import ProcesadorRegistroRetributivo
    from procesar_datos_triodos import ProcesadorTriodos
    from generar_informe_optimizado import GeneradorInformeOptimizado
    from validador_mapeo import ValidadorMapeoGeneral, ValidadorMapeoTriodos
except ImportError as e:
    st.error(f"Error al importar módulos: {e}")
    st.info("Asegúrate de que todos los archivos necesarios estén en la carpeta 04_SCRIPTS")
    st.stop()

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


def crear_carpetas_necesarias():
    """Crea las carpetas necesarias solo para desarrollo local"""
    # Solo crear carpetas si NO estamos en producción (Streamlit Cloud)
    # En producción, todo se maneja en memoria sin tocar disco
    if os.getenv('STREAMLIT_SHARING_MODE') or os.getenv('STREAMLIT_CLOUD'):
        return  # No crear carpetas en producción

    base_path = Path(__file__).parent
    carpetas = [
        base_path / "01_DATOS_SIN_PROCESAR",
        base_path / "02_RESULTADOS",
        base_path / "03_LOGS",
        base_path / "05_INFORMES"
    ]
    for carpeta in carpetas:
        carpeta.mkdir(exist_ok=True)


def detectar_campos_faltantes(archivo_bytes, tipo_archivo, password=None):
    """
    Solo detecta campos faltantes, NO muestra UI ni aplica mapeo.

    Returns:
        Tuple (validador, info_deteccion, error_mensaje)
        info_deteccion contiene toda la info sobre campos faltantes
    """
    try:
        # Guardar temporalmente el archivo para análisis
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(archivo_bytes)
            tmp_path = Path(tmp_file.name)

        try:
            # Crear validador según el tipo
            if tipo_archivo == "Triodos":
                validador = ValidadorMapeoTriodos()
                # Si es Triodos, necesitamos desencriptar primero
                if password:
                    from procesar_datos_triodos import ProcesadorTriodos
                    procesador_temp = ProcesadorTriodos()
                    procesador_temp.password = password
                    archivo_decrypted = procesador_temp.abrir_archivo_protegido(tmp_path)
                    excel_file = pd.ExcelFile(archivo_decrypted)
                else:
                    excel_file = pd.ExcelFile(tmp_path)
            else:
                validador = ValidadorMapeoGeneral()
                excel_file = pd.ExcelFile(tmp_path)

            # PASO 1: Validar hojas (SOLO detectar, NO mostrar UI)
            resultado_hojas = validador.validar_hojas(excel_file, validador.hojas_requeridas)

            # PASO 2: Validar variables de la hoja principal (BASE GENERAL)
            nombre_hoja_principal = validador.obtener_nombre_hoja('BASE GENERAL')

            try:
                if tipo_archivo == "Triodos" and password:
                    # Reabrir el archivo desencriptado
                    archivo_decrypted.seek(0)
                    df = pd.read_excel(archivo_decrypted, sheet_name=nombre_hoja_principal)
                else:
                    df = pd.read_excel(tmp_path, sheet_name=nombre_hoja_principal)

                # Limpiar nombres de columnas
                df.columns = df.columns.str.strip()

            except Exception as e:
                return None, None, f"Error al leer la hoja '{nombre_hoja_principal}': {str(e)}"

            resultado_variables = validador.validar_variables(df, validador.variables_criticas)

            # PASO 3: Validar variables de hojas de complementos (SOLO detectar)
            resultados_complementos = {}
            for nombre_hoja_config in ['COMPLEMENTOS SALARIALES', 'COMPLEMENTOS EXTRASALARIALES']:
                nombre_hoja_real = validador.obtener_nombre_hoja(nombre_hoja_config)

                if nombre_hoja_real not in excel_file.sheet_names:
                    continue

                try:
                    if tipo_archivo == "Triodos" and password:
                        archivo_decrypted.seek(0)
                        df_comp = pd.read_excel(archivo_decrypted, sheet_name=nombre_hoja_real)
                    else:
                        df_comp = pd.read_excel(tmp_path, sheet_name=nombre_hoja_real)

                    df_comp.columns = df_comp.columns.str.strip()
                    resultado_vars_comp = validador.validar_variables(df_comp, validador.columnas_config_complementos)
                    resultados_complementos[nombre_hoja_config] = resultado_vars_comp

                except Exception as e:
                    pass

            # Retornar validador e información de detección
            info_deteccion = {
                'hojas': resultado_hojas,
                'variables': resultado_variables,
                'complementos': resultados_complementos
            }

            return validador, info_deteccion, None

        finally:
            # Limpiar archivo temporal
            if tmp_path.exists():
                tmp_path.unlink()

    except Exception as e:
        return None, None, f"Error durante la validación: {str(e)}"


def main():
    # Crear carpetas necesarias al inicio
    crear_carpetas_necesarias()
    
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

    st.markdown("---")

    # Sidebar con información
    with st.sidebar:
        st.header("ℹ️ Información")

        st.markdown("""
        ### Cómo usar:
        1. Selecciona el tipo de archivo
        2. Selecciona la acción deseada
        3. Sube tu archivo Excel
        4. Procesa y descarga resultados

        ### Acciones disponibles:
        - **Ambas**: Procesa datos Y genera informe (recomendado)
        - **Procesar Datos**: Solo genera Excel procesado
        - **Generar Informe**: Genera informe Word de archivo ya procesado

        ### Formatos admitidos:
        - **General**: Archivo con hoja "BASE GENERAL" (usa `procesar_datos.py`)
        - **Triodos**: Archivo Triodos Bank protegido (usa `procesar_datos_triodos.py`)

        ### Archivos generados:
        - **Excel**: Datos procesados con columnas equiparadas
        - **Word**: Informe completo con gráficos y tablas (usa `generar_informe_optimizado.py`)
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
                value="",
                type="password",
                help="Introduce la contraseña para desbloquear el archivo Excel",
                placeholder="Introduce la contraseña..."
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
        # Limpiar validación anterior si se sube un nuevo archivo
        if 'ultimo_archivo' in st.session_state and st.session_state['ultimo_archivo'] != archivo_subido.name:
            st.session_state.pop('validacion_completa', None)
            st.session_state.pop('validador', None)

        st.session_state['ultimo_archivo'] = archivo_subido.name
        st.success(f"✅ Archivo cargado: **{archivo_subido.name}** ({archivo_subido.size / 1024:.2f} KB)")

        # Validar tamaño (50MB máximo)
        if archivo_subido.size > 50 * 1024 * 1024:
            st.error("❌ El archivo es demasiado grande. Tamaño máximo: 50MB")
            return

        st.markdown("---")

        # Botón de procesamiento
        st.header("3️⃣ Validar y Procesar Datos")

        # PASO 1: Validación y mapeo (si no se ha hecho)
        if 'validacion_completa' not in st.session_state:
            # Sub-paso: Iniciar detección
            if 'info_deteccion' not in st.session_state:
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🔍 Validar Archivo y Mapear Campos", type="secondary", use_container_width=True):
                        # Validar que se haya introducido contraseña si el archivo está protegido
                        if archivo_protegido and not password:
                            st.error("❌ Por favor, introduce la contraseña del archivo antes de validar.")
                            st.stop()

                        # Leer archivo como bytes
                        archivo_bytes = archivo_subido.read()
                        archivo_subido.seek(0)  # Reset para poder leerlo después

                        # SOLO detectar, NO mostrar UI
                        validador, info_deteccion, error = detectar_campos_faltantes(archivo_bytes, tipo_archivo, password)

                        if error:
                            st.error(f"❌ Error: {error}")
                            st.stop()

                        # Guardar en session_state para mostrar UI después
                        st.session_state['validador_temp'] = validador
                        st.session_state['info_deteccion'] = info_deteccion
                        st.rerun()

            # Sub-paso: Mostrar UI de mapeo si hay info de detección
            if 'info_deteccion' in st.session_state:
                info = st.session_state['info_deteccion']
                validador = st.session_state['validador_temp']

                resultado_hojas = info['hojas']
                resultado_variables = info['variables']
                resultados_complementos = info['complementos']

                # UI para hojas faltantes
                if resultado_hojas['faltantes']:
                    st.warning(f"⚠️ Se encontraron {len(resultado_hojas['faltantes'])} hojas faltantes")
                    st.markdown("### 🗂️ Mapeo de Hojas")

                    for hoja_faltante in resultado_hojas['faltantes']:
                        st.markdown(f"**Hoja esperada:** `{hoja_faltante}` ⚠️ **OBLIGATORIA**")
                        st.selectbox(
                            f"Selecciona la hoja real para '{hoja_faltante}':",
                            options=["[Selecciona una opción]"] + resultado_hojas['disponibles'],
                            key=f"mapeo_hoja_{hoja_faltante}"
                        )
                        st.markdown("---")

                # UI para variables faltantes
                if resultado_variables['faltantes']:
                    st.warning(f"⚠️ Se encontraron {len(resultado_variables['faltantes'])} variables faltantes")
                    st.markdown("### 📊 Mapeo de Variables")

                    for clave_interna, nombre_esperado in resultado_variables['faltantes'].items():
                        es_obligatoria = clave_interna in validador.variables_obligatorias
                        etiqueta = " ⚠️ **OBLIGATORIA**" if es_obligatoria else " (opcional)"
                        st.markdown(f"**Variable esperada:** `{nombre_esperado}`{etiqueta}")

                        if es_obligatoria:
                            opciones = ["[Selecciona una opción]"] + resultado_variables['disponibles']
                        else:
                            opciones = ["⛔ No mapear (omitir)"] + resultado_variables['disponibles']

                        st.selectbox(
                            f"Selecciona la columna real para '{nombre_esperado}':",
                            options=opciones,
                            key=f"mapeo_var_{clave_interna}"
                        )
                        st.markdown("---")

                # Botón para confirmar mapeo
                st.markdown("---")
                st.info("⚠️ Revisa las selecciones arriba. Cuando hayas terminado, presiona 'Confirmar Mapeo'.")

                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("✅ Confirmar Mapeo", type="primary", use_container_width=True):
                        # Aplicar mapeos desde session_state
                        # Hojas
                        if resultado_hojas['faltantes']:
                            mapeo_hojas = {}
                            for hoja in resultado_hojas['faltantes']:
                                sel = st.session_state.get(f"mapeo_hoja_{hoja}")
                                if sel and sel != "[Selecciona una opción]":
                                    mapeo_hojas[hoja] = sel

                            if len(mapeo_hojas) < len(resultado_hojas['faltantes']):
                                st.error("⚠️ Debes mapear todas las hojas obligatorias")
                                st.stop()

                            validador.aplicar_mapeo_hojas(mapeo_hojas)

                        # Variables
                        if resultado_variables['faltantes']:
                            mapeo_vars = {}
                            for clave in resultado_variables['faltantes'].keys():
                                sel = st.session_state.get(f"mapeo_var_{clave}")
                                if sel and sel not in ["[Selecciona una opción]", "⛔ No mapear (omitir)"]:
                                    mapeo_vars[clave] = sel

                            # Verificar obligatorias
                            obligatorias_faltantes = [
                                clave for clave in validador.variables_obligatorias
                                if clave in resultado_variables['faltantes'] and clave not in mapeo_vars
                            ]
                            if obligatorias_faltantes:
                                nombres = [validador.variables_criticas[c] for c in obligatorias_faltantes]
                                st.error(f"⚠️ Debes mapear las variables obligatorias: {', '.join(nombres)}")
                                st.stop()

                            validador.aplicar_mapeo_variables(mapeo_vars)

                        # Guardar y continuar
                        st.session_state['validador'] = validador
                        st.session_state['validacion_completa'] = True
                        st.session_state.pop('info_deteccion', None)
                        st.session_state.pop('validador_temp', None)
                        st.success("✅ Mapeo confirmado!")
                        st.rerun()

        # PASO 2: Procesamiento (solo si la validación está completa)
        if st.session_state.get('validacion_completa', False):
            st.success("✅ Archivo validado. Campos mapeados correctamente.")

            # Mostrar resumen del mapeo
            if 'validador' in st.session_state:
                validador = st.session_state['validador']

                with st.expander("📋 Ver mapeo aplicado", expanded=False):
                    if validador.mapeo_hojas:
                        st.markdown("**Hojas mapeadas:**")
                        for esperado, real in validador.mapeo_hojas.items():
                            st.markdown(f"- `{esperado}` → `{real}`")

                    if validador.mapeo_variables:
                        st.markdown("**Variables mapeadas:**")
                        for clave, real in validador.mapeo_variables.items():
                            nombre_esperado = validador.variables_criticas.get(clave, clave)
                            st.markdown(f"- `{nombre_esperado}` → `{real}`")

            # Botón para reiniciar validación
            col_reset1, col_reset2, col_reset3 = st.columns([1, 2, 1])
            with col_reset2:
                if st.button("🔄 Reiniciar Validación", type="secondary", use_container_width=True, help="Volver a validar y mapear campos"):
                    st.session_state.pop('validacion_completa', None)
                    st.session_state.pop('validador', None)
                    st.rerun()

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                # Cambiar el texto del botón según la acción
                texto_boton = {
                    "Ambas": "🚀 Procesar y Generar Informe",
                    "Procesar Datos": "📊 Procesar Datos",
                    "Generar Informe": "📄 Generar Informe"
                }

                if st.button(texto_boton[accion], type="primary", use_container_width=True):
                    # Validar que se haya introducido contraseña si el archivo está protegido
                    if archivo_protegido and not password:
                        st.error("❌ Por favor, introduce la contraseña del archivo antes de procesar.")
                        st.stop()

                    with st.spinner(f"{accion}... Esto puede tardar unos segundos."):
                        try:
                            excel_procesado = None
                            informe_word = None

                            # PASO 1: Procesar datos (si corresponde)
                            if accion in ["Ambas", "Procesar Datos"]:
                                with st.spinner("📊 Procesando datos..."):
                                    # Guardar temporalmente el archivo (se borra automáticamente)
                                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                                        tmp_file.write(archivo_bytes)
                                        tmp_path = Path(tmp_file.name)

                                    try:
                                        # Seleccionar procesador según tipo
                                        # Pasar el validador al procesador si existe
                                        validador_a_usar = st.session_state.get('validador', None)

                                        if tipo_archivo == "Triodos":
                                            st.info("📋 Usando procesador de Triodos Bank...")
                                            procesador = ProcesadorTriodos(validador=validador_a_usar)
                                            # Configurar la contraseña si se proporcionó
                                            if password:
                                                procesador.password = password
                                        else:
                                            st.info("📋 Usando procesador general...")
                                            procesador = ProcesadorRegistroRetributivo(validador=validador_a_usar)

                                        # Procesar el archivo
                                        resultado = procesador.procesar_archivo(tmp_path)

                                        if resultado['estado'] == 'ÉXITO':
                                            # Leer el archivo generado desde disco
                                            carpeta_resultados = Path(__file__).parent / "02_RESULTADOS"
                                            archivo_resultado = carpeta_resultados / resultado['archivo_resultado']

                                            if archivo_resultado.exists():
                                                # Cargar en memoria
                                                with open(archivo_resultado, 'rb') as f:
                                                    excel_procesado = io.BytesIO(f.read())

                                                # IMPORTANTE: Borrar archivo del disco inmediatamente
                                                try:
                                                    archivo_resultado.unlink()
                                                    st.info("🔒 Archivo temporal eliminado del servidor")
                                                except Exception as e:
                                                    st.warning(f"⚠️ No se pudo eliminar archivo temporal: {e}")

                                                # Guardar estadísticas
                                                st.session_state['estadisticas'] = {
                                                    'total_registros': resultado['registros_procesados'],
                                                    'tiempo_procesamiento': resultado['tiempo_procesamiento']
                                                }

                                                st.success(f"✅ Datos procesados: {resultado['registros_procesados']} registros en {resultado['tiempo_procesamiento']:.1f}s")
                                            else:
                                                raise Exception(f"No se encontró el archivo resultado: {archivo_resultado}")
                                        else:
                                            raise Exception(resultado.get('error', 'Error desconocido en el procesamiento'))

                                    finally:
                                        # Limpiar archivo temporal de entrada
                                        if tmp_path.exists():
                                            tmp_path.unlink()

                            # PASO 2: Generar informe (si corresponde)
                            if accion in ["Ambas", "Generar Informe"]:
                                with st.spinner("📄 Generando informe Word..."):
                                    # Determinar qué archivo usar para el informe
                                    archivo_para_informe = None

                                    if accion == "Ambas":
                                        # Usar el archivo recién procesado
                                        if excel_procesado is None:
                                            raise Exception("Error: No se pudo procesar el archivo Excel")

                                        st.info("📋 Generando informe con datos recién procesados...")
                                        # Guardar temporalmente el archivo procesado
                                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp_file:
                                            excel_procesado.seek(0)
                                            tmp_file.write(excel_procesado.read())
                                            archivo_para_informe = Path(tmp_file.name)

                                    elif accion == "Generar Informe":
                                        # Usar el archivo original subido (ya procesado previamente)
                                        st.info("📋 Generando informe directamente del archivo subido...")
                                        # Guardar temporalmente el archivo original
                                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp_file:
                                            tmp_file.write(archivo_bytes)
                                            archivo_para_informe = Path(tmp_file.name)

                                    try:
                                        # Crear generador de informes
                                        generador = GeneradorInformeOptimizado()

                                        # Cargar datos desde el archivo
                                        try:
                                            generador.df = pd.read_excel(archivo_para_informe)
                                        except Exception as e:
                                            raise Exception(f"Error al leer el archivo Excel: {str(e)}. Asegúrate de que el archivo tenga el formato correcto.")

                                        # Mapear valores de la columna Sexo
                                        if 'Sexo' in generador.df.columns:
                                            generador.df['Sexo'] = generador.df['Sexo'].map({
                                                'Hombres': 'H',
                                                'Mujeres': 'M'
                                            }).fillna(generador.df['Sexo'])

                                        st.info(f"📊 Datos cargados: {len(generador.df)} registros")

                                        # Generar el informe (tipo CONSOLIDADO por defecto)
                                        if generador.generar_informe('CONSOLIDADO'):
                                            # Leer el archivo Word generado
                                            carpeta_informes = Path(__file__).parent / "05_INFORMES"

                                            # Buscar el archivo más reciente
                                            if carpeta_informes.exists():
                                                archivos_word = list(carpeta_informes.glob('registro_retributivo_*.docx'))
                                                if archivos_word:
                                                    archivo_word = max(archivos_word, key=lambda x: x.stat().st_mtime)

                                                    # Cargar en memoria
                                                    with open(archivo_word, 'rb') as f:
                                                        informe_word = io.BytesIO(f.read())

                                                    # IMPORTANTE: Borrar archivo del disco inmediatamente
                                                    try:
                                                        archivo_word.unlink()
                                                        st.info("🔒 Informe temporal eliminado del servidor")
                                                    except Exception as e:
                                                        st.warning(f"⚠️ No se pudo eliminar informe temporal: {e}")

                                                    # Limpiar imágenes temporales de gráficos
                                                    try:
                                                        for temp_file in generador.archivos_temp:
                                                            if os.path.exists(temp_file):
                                                                os.remove(temp_file)
                                                    except:
                                                        pass

                                                    st.success("✅ Informe generado correctamente")
                                                else:
                                                    raise Exception("No se encontró el archivo Word generado")
                                            else:
                                                raise Exception("No se pudo crear el informe")
                                        else:
                                            raise Exception("Error al generar el informe")

                                    finally:
                                        # Limpiar archivo temporal
                                        if archivo_para_informe and archivo_para_informe.exists():
                                            archivo_para_informe.unlink()

                            # Guardar en session_state
                            if excel_procesado:
                                excel_procesado.seek(0)
                                st.session_state['archivo_procesado'] = excel_procesado

                            if informe_word:
                                st.session_state['informe_word'] = informe_word

                            st.session_state['nombre_archivo'] = archivo_subido.name.replace('.xlsx', '').replace('.xls', '')
                            st.session_state['accion_realizada'] = accion

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
                        • Tiempo de procesamiento: <strong>{:.2f}s</strong>
                    </div>
                """.format(
                    stats['total_registros'],
                    stats.get('tiempo_procesamiento', 0)
                ), unsafe_allow_html=True)

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

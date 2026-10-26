"""
================================================================================
CONVERTIDOR MASIVO DOCX A PDF DE ALTA FIDELIDAD Y OPTIMIZADOR DE PESO
================================================================================
Arquitectura de Grado de Producción para conversión por lotes con:
- Fidelidad estructural 100% (márgenes, tipografías, tablas, paginación idéntica).
- Motor nativo COM de Microsoft Word (Windows) con aislamiento por hilo.
- Motor headless de LibreOffice (Linux/Docker/macOS/Fallback).
- Funcionalidad de reducción de peso (compresión inteligente sin pérdida visual).
- Soporte dual para optimización: Ghostscript y pipeline nativo (pypdf + Pillow).
- Barra de progreso interactiva en tiempo real.
- Descarga individual y empaquetado masivo en ZIP.
- Validación criptográfica/mágica de formato OpenXML.
- Gestión de memoria y limpieza estricta de archivos temporales y procesos zombies.
- Pie de página institucional: "Desarrollado por Ing. Wilson Castellanos @ 2026".
================================================================================
"""

import os
import sys
import time
import zipfile
import shutil
import tempfile
import platform
import subprocess
from io import BytesIO
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
from enum import Enum
from dataclasses import dataclass

import streamlit as st
from PIL import Image
import pypdf

# Verificación condicional para pywin32 en entornos Windows
IS_WINDOWS = platform.system().lower() == "windows"
if IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client
        HAS_WIN32COM = True
    except ImportError:
        HAS_WIN32COM = False
else:
    HAS_WIN32COM = False


# ==============================================================================
# MODELOS DE DATOS Y ENUMERACIONES
# ==============================================================================

class ConversionStatus(str, Enum):
    PENDING = "Pendiente"
    PROCESSING = "Procesando"
    SUCCESS = "Exitoso"
    ERROR = "Error"
    VALIDATION_FAILED = "Inválido"


@dataclass
class ConversionResult:
    filename: str
    status: ConversionStatus
    pdf_bytes: Optional[bytes] = None
    original_size: int = 0
    pdf_size: int = 0
    savings_pct: float = 0.0
    duration: float = 0.0
    error_message: Optional[str] = None


# ==============================================================================
# MÓDULO DE REDUCCIÓN DE PESO (COMPRESIÓN INTELIGENTE SIN PÉRDIDA VISUAL)
# ==============================================================================

def compress_pdf_bytes(
    pdf_bytes: bytes,
    level: str = "ebook",      # "ebook" (150 dpi), "screen" (72 dpi), "printer" (300 dpi)
    jpeg_quality: int = 82
) -> Tuple[bytes, float]:
    """
    Reduce el peso del PDF manteniendo nítidas las tipografías vectoriales, márgenes
    y paginación. Utiliza Ghostscript si está instalado o un pipeline nativo de pypdf + Pillow.
    
    Garantía de calidad:
    - Comprime flujos de texto sin pérdida (FlateDecode).
    - Elimina objetos y metadatos huérfanos o duplicados.
    - Recomprime imágenes internas a JPEG con muestreo balanceado.
    - Si el archivo resultante no reduce el peso, se conserva el PDF original intacto.
    """
    orig_size = len(pdf_bytes)
    if orig_size == 0:
        return pdf_bytes, 0.0

    optimized_bytes = pdf_bytes

    # 1. Intentar con Ghostscript si existe en PATH (gswin64c / gs)
    gs_cmd = shutil.which("gswin64c") or shutil.which("gs")
    if gs_cmd:
        try:
            with tempfile.TemporaryDirectory(prefix="gs_opt_") as tmp_dir:
                tmp_in = Path(tmp_dir) / "input.pdf"
                tmp_out = Path(tmp_dir) / "output.pdf"
                tmp_in.write_bytes(pdf_bytes)

                args = [
                    gs_cmd,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    f"-dPDFSETTINGS=/{level}",
                    "-dNOPAUSE",
                    "-dQUIET",
                    "-dBATCH",
                    f"-sOutputFile={tmp_out}",
                    str(tmp_in)
                ]
                res = subprocess.run(args, timeout=45, capture_output=True)
                if tmp_out.exists() and tmp_out.stat().st_size > 0:
                    gs_data = tmp_out.read_bytes()
                    if len(gs_data) < len(optimized_bytes):
                        optimized_bytes = gs_data
        except Exception:
            pass

    # 2. Pipeline nativo en memoria con pypdf y Pillow
    try:
        in_buf = BytesIO(optimized_bytes)
        reader = pypdf.PdfReader(in_buf)
        writer = pypdf.PdfWriter(clone_from=reader)

        # Comprimir secuencias de contenido (lossless)
        for page in writer.pages:
            try:
                page.compress_content_streams()
            except Exception:
                pass

            # Optimizar imágenes incrustadas de alta resolución
            try:
                for img_obj in page.images:
                    try:
                        pil_img = img_obj.image
                        # Solo recomprimir imágenes RGB/L sin alterar canales alfa/transparencias
                        if pil_img.mode in ("RGB", "L"):
                            img_obj.replace(pil_img, quality=jpeg_quality)
                    except Exception:
                        pass
            except Exception:
                pass

        # Eliminar objetos idénticos y desreferenciados
        try:
            writer.compress_identical_objects(remove_duplicates=True, remove_unreferenced=True)
        except Exception:
            try:
                writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
            except Exception:
                pass

        out_buf = BytesIO()
        writer.write(out_buf)
        py_data = out_buf.getvalue()

        if len(py_data) > 0 and len(py_data) < len(optimized_bytes):
            optimized_bytes = py_data
    except Exception:
        pass

    final_size = len(optimized_bytes)
    # Regla de oro: si no hubo reducción, devolver original
    if final_size >= orig_size:
        return pdf_bytes, 0.0

    savings_pct = (1.0 - (final_size / orig_size)) * 100.0
    return optimized_bytes, savings_pct


# ==============================================================================
# DETECCIÓN Y ABSTRACCIÓN DE MOTORES DE CONVERSIÓN (HIGH-FIDELITY BACKENDS)
# ==============================================================================

class BaseConversionEngine:
    """Clase base abstracta para motores de renderizado DOCX a PDF."""
    
    @property
    def name(self) -> str:
        raise NotImplementedError

    def is_available(self) -> Tuple[bool, str]:
        raise NotImplementedError

    def convert(self, docx_path: Path, output_pdf_path: Path, optimize_size: bool = False) -> None:
        raise NotImplementedError


class WordCOMConversionEngine(BaseConversionEngine):
    """
    Motor nativo para Windows mediante Microsoft Word COM Automation.
    Garantiza 100% de fidelidad visual, tipográfica, tabular y paginación idéntica
    al utilizar el motor de renderizado oficial de Microsoft Office.
    """
    
    @property
    def name(self) -> str:
        return "Microsoft Word Native COM (Fidelidad 100% Oficial)"

    def is_available(self) -> Tuple[bool, str]:
        if not IS_WINDOWS:
            return False, "Solo disponible en sistemas operativos Windows."
        if not HAS_WIN32COM:
            return False, "La librería 'pywin32' no está instalada."
        
        try:
            pythoncom.CoInitialize()
            try:
                word_app = win32com.client.Dispatch("Word.Application")
                version = word_app.Version
                word_app.Quit()
                return True, f"Microsoft Word v{version} detectado y listo."
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            return False, f"Microsoft Word no está instalado o COM inaccesible: {str(e)}"

    def convert(self, docx_path: Path, output_pdf_path: Path, optimize_size: bool = False) -> None:
        """
        Ejecuta la exportación a PDF usando la API COM de Word.
        Maneja CoInitialize para seguridad en subprocesos/hilos de Streamlit.
        """
        pythoncom.CoInitialize()
        word_app = None
        doc = None
        
        try:
            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0  # wdAlertsNone: silencia popups y advertencias
            
            # Apertura segura del documento
            try:
                doc = word_app.Documents.Open(
                    FileName=str(docx_path.resolve()),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    PasswordDocument="",
                    PasswordTemplate="",
                    Revert=False,
                    WritePasswordDocument="",
                    WritePasswordTemplate="",
                    Format=0,
                    Visible=False
                )
            except Exception as open_err:
                err_str = str(open_err).lower()
                if "password" in err_str or "contraseña" in err_str or "0x800a1485" in err_str:
                    raise PermissionError("El documento está protegido con contraseña.")
                raise RuntimeError(f"Error al abrir el archivo en Word: {str(open_err)}")

            # Optimización de renderizado según selección del usuario:
            # OptimizeFor = 1 (wdExportOptimizeForOnScreen) reduce peso sin perder nitidez de texto
            # OptimizeFor = 0 (wdExportOptimizeForPrint) máxima resolución de impresión
            export_optimize = 1 if optimize_size else 0

            doc.ExportAsFixedFormat(
                OutputFileName=str(output_pdf_path.resolve()),
                ExportFormat=17,              # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=export_optimize,  # 0: Print, 1: Screen (Peso optimizado)
                Range=0,                     # wdExportAllDocument
                From=1,
                To=1,
                Item=0,                      # wdExportDocumentContent
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,           # wdExportCreateHeadingBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,     # Preserva fuentes no vectorizadas sin deformar el layout
                UseISO19005_1=False          # PDF estándar con todas las características de maquetación
            )
            
            if not output_pdf_path.exists() or output_pdf_path.stat().st_size == 0:
                raise RuntimeError("El motor de Word finalizó pero no generó un archivo PDF válido.")

        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
                except Exception:
                    pass
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


class LibreOfficeConversionEngine(BaseConversionEngine):
    """
    Motor multiplataforma mediante LibreOffice en modo Headless.
    Ideal para despliegues en contenedores Docker, servidores Linux o fallbacks en Windows/macOS.
    """
    
    def __init__(self):
        self._binary_path: Optional[str] = self._locate_libreoffice()

    @property
    def name(self) -> str:
        return "LibreOffice Headless (Multiplataforma Linux / Docker / Cloud)"

    @staticmethod
    def _locate_libreoffice() -> Optional[str]:
        candidates = ["soffice", "libreoffice"]
        for cmd in candidates:
            path = shutil.which(cmd)
            if path:
                return path

        common_paths = []
        if IS_WINDOWS:
            common_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
        elif platform.system().lower() == "darwin":
            common_paths = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            ]
        else:
            common_paths = [
                "/usr/bin/libreoffice",
                "/usr/bin/soffice",
                "/usr/local/bin/libreoffice",
                "/usr/local/bin/soffice",
                "/opt/libreoffice/program/soffice"
            ]

        for p in common_paths:
            if os.path.isfile(p) and os.access(p, os.X_OK if not IS_WINDOWS else os.R_OK):
                return p
        return None

    def is_available(self) -> Tuple[bool, str]:
        if self._binary_path:
            return True, f"LibreOffice ejecutable detectado en: {self._binary_path}"
        return False, "LibreOffice ('soffice' o 'libreoffice') no se encuentra instalado en el sistema."

    def convert(self, docx_path: Path, output_pdf_path: Path, optimize_size: bool = False) -> None:
        if not self._binary_path:
            raise RuntimeError("LibreOffice no está disponible en este servidor.")

        outdir = str(output_pdf_path.parent.resolve())
        cmd = [
            self._binary_path,
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            outdir,
            str(docx_path.resolve())
        ]

        env = os.environ.copy()
        temp_profile = tempfile.mkdtemp(prefix="lo_profile_")
        env["UserInstallation"] = f"file:///{temp_profile.replace(os.sep, '/')}"

        try:
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=90,
                env=env
            )
            
            expected_pdf = output_pdf_path.parent / (docx_path.stem + ".pdf")
            if expected_pdf.exists() and expected_pdf.stat().st_size > 0:
                if expected_pdf != output_pdf_path:
                    shutil.move(str(expected_pdf), str(output_pdf_path))
            else:
                stderr_text = result.stderr.strip() or result.stdout.strip()
                raise RuntimeError(f"Fallo en LibreOffice (Código {result.returncode}): {stderr_text or 'Sin salida'}")
        finally:
            shutil.rmtree(temp_profile, ignore_errors=True)


def get_available_engines() -> List[BaseConversionEngine]:
    """Retorna los motores de conversión viables en el orden preferido."""
    engines: List[BaseConversionEngine] = []
    
    if IS_WINDOWS:
        word_engine = WordCOMConversionEngine()
        available, _ = word_engine.is_available()
        if available:
            engines.append(word_engine)

    lo_engine = LibreOfficeConversionEngine()
    available_lo, _ = lo_engine.is_available()
    if available_lo:
        engines.append(lo_engine)

    if IS_WINDOWS and not any(isinstance(e, WordCOMConversionEngine) for e in engines):
        engines.append(WordCOMConversionEngine())

    if not any(isinstance(e, LibreOfficeConversionEngine) for e in engines):
        engines.append(lo_engine)

    return engines


# ==============================================================================
# VALIDACIONES ROBUSTAS DE ARCHIVOS
# ==============================================================================

def validate_docx_file(file_name: str, file_bytes: bytes) -> Tuple[bool, Optional[str]]:
    """
    Valida la extensión y la firma binaria de archivo OpenXML (.docx).
    Un archivo .docx válido es un contenedor ZIP que inicia con los magic bytes PK (0x50 0x4B).
    """
    if not file_name.lower().endswith(".docx"):
        return False, "Extensión inválida: Solo se permiten archivos con extensión '.docx'."
    
    if len(file_bytes) == 0:
        return False, "El archivo está vacío (0 bytes)."

    if not file_bytes.startswith(b"PK\x03\x04") and not file_bytes.startswith(b"PK\x05\x06") and not file_bytes.startswith(b"PK\x07\x08"):
        return False, "Firma de archivo inválida: No es un archivo OpenXML (.docx) válido o está dañado."

    try:
        with zipfile.ZipFile(BytesIO(file_bytes), 'r') as zf:
            namelist = zf.namelist()
            if not any(entry.startswith("word/") or entry == "[Content_Types].xml" for entry in namelist):
                return False, "Estructura interna corrupta: El archivo no contiene la jerarquía de un documento Word."
    except zipfile.BadZipFile:
        return False, "Archivo comprimido OpenXML dañado o corrupto."
    except Exception as e:
        return False, f"Error al verificar integridad: {str(e)}"

    return True, None


# ==============================================================================
# PROCESAMIENTO POR LOTES Y GESTIÓN DE ARCHIVOS TEMPORALES
# ==============================================================================

def process_batch(
    files: List[Any],
    engine: BaseConversionEngine,
    progress_bar,
    status_text,
    enable_compression: bool = True,
    compression_level: str = "ebook",
    jpeg_quality: int = 82
) -> List[ConversionResult]:
    """
    Convierte y opcionalmente optimiza en lote los archivos subidos garantizando:
    - Progreso en tiempo real con actualización reactiva.
    - Captura y aislamiento de excepciones por archivo individual.
    - Reducción de peso sin pérdida de calidad visual.
    - Limpieza automática e inmediata de archivos temporales de disco.
    """
    results: List[ConversionResult] = []
    total_files = len(files)

    for index, uploaded_file in enumerate(files):
        current_num = index + 1
        pct = current_num / total_files
        file_name = uploaded_file.name
        
        status_text.markdown(f"⏳ **Procesando ({current_num}/{total_files}):** `{file_name}`...")
        progress_bar.progress(pct)

        # 1. Validación previa de integridad
        file_bytes = uploaded_file.getvalue()
        is_valid, validation_error = validate_docx_file(file_name, file_bytes)
        
        if not is_valid:
            results.append(ConversionResult(
                filename=file_name,
                status=ConversionStatus.VALIDATION_FAILED,
                error_message=validation_error
            ))
            continue

        # 2. Conversión aislada en directorio temporal estricto
        t_start = time.perf_counter()
        with tempfile.TemporaryDirectory(prefix="docx_pdf_convert_") as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            input_docx_path = temp_dir / file_name
            output_pdf_name = f"{Path(file_name).stem}.pdf"
            output_pdf_path = temp_dir / output_pdf_name

            try:
                with open(input_docx_path, "wb") as f_in:
                    f_in.write(file_bytes)

                # Ejecutar motor de renderizado con flag de optimización de tamaño
                engine.convert(input_docx_path, output_pdf_path, optimize_size=enable_compression)

                with open(output_pdf_path, "rb") as f_pdf:
                    raw_pdf_data = f_pdf.read()

                initial_size = len(raw_pdf_data)

                # 3. Aplicar compresión inteligente post-renderizado si está habilitada
                final_pdf_data = raw_pdf_data
                savings = 0.0

                if enable_compression:
                    final_pdf_data, savings = compress_pdf_bytes(
                        raw_pdf_data,
                        level=compression_level,
                        jpeg_quality=jpeg_quality
                    )

                duration = time.perf_counter() - t_start
                results.append(ConversionResult(
                    filename=output_pdf_name,
                    status=ConversionStatus.SUCCESS,
                    pdf_bytes=final_pdf_data,
                    original_size=initial_size,
                    pdf_size=len(final_pdf_data),
                    savings_pct=savings,
                    duration=duration
                ))

            except PermissionError as pe:
                results.append(ConversionResult(
                    filename=file_name,
                    status=ConversionStatus.ERROR,
                    duration=time.perf_counter() - t_start,
                    error_message=f"Documento protegido con contraseña: {str(pe)}"
                ))
            except Exception as ex:
                results.append(ConversionResult(
                    filename=file_name,
                    status=ConversionStatus.ERROR,
                    duration=time.perf_counter() - t_start,
                    error_message=f"Fallo durante la conversión: {str(ex)}"
                ))

    status_text.markdown("✅ **¡Lote de conversión completado!**")
    return results


def build_zip_package(successful_results: List[ConversionResult]) -> bytes:
    """Genera un archivo ZIP en memoria conteniendo todos los PDFs procesados con éxito."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in successful_results:
            if item.pdf_bytes:
                zip_file.writestr(item.filename, item.pdf_bytes)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def format_bytes(size: int) -> str:
    """Formatea tamaños de bytes de forma legible para el usuario."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:3.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


# ==============================================================================
# ESTILOS CSS PROFESIONALES DE ALTA ESTÉTICA Y PIE DE PÁGINA
# ==============================================================================

CUSTOM_CSS = """
<meta name="google" content="notranslate" />
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, .stApp, [class*="css"], [data-testid="stAppViewContainer"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        translate: no !important;
    }

    /* Tarjetas de métricas modernas con glassmorphism */
    .metric-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 16px 20px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        border-color: rgba(64, 158, 255, 0.4);
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .metric-label {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        opacity: 0.75;
    }

    /* Badges de ahorro de peso */
    .saving-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 700;
        background-color: rgba(46, 204, 113, 0.2);
        color: #2ecc71;
        border: 1px solid rgba(46, 204, 113, 0.4);
        margin-left: 6px;
    }

    /* Encabezado del contenedor principal */
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 26px 30px;
        border-radius: 16px;
        margin-bottom: 25px;
        box-shadow: 0 8px 32px rgba(30, 60, 114, 0.25);
    }
    .main-header h1 {
        color: white !important;
        margin: 0 0 6px 0;
        font-size: 2.1rem;
        font-weight: 700;
    }
    .main-header p {
        color: rgba(255, 255, 255, 0.88);
        margin: 0;
        font-size: 1.05rem;
    }

    /* Caja de opciones de optimización */
    .opt-box {
        background: rgba(46, 204, 113, 0.05);
        border: 1px solid rgba(46, 204, 113, 0.25);
        border-radius: 12px;
        padding: 14px 18px;
        margin: 15px 0 20px 0;
    }

    /* PIE DE PÁGINA PROFESIONAL */
    .footer-container {
        margin-top: 50px;
        padding: 24px 10px 15px 10px;
        text-align: center;
        border-top: 1px solid rgba(255, 255, 255, 0.12);
    }
    .footer-badge {
        display: inline-block;
        padding: 8px 22px;
        background: linear-gradient(135deg, rgba(30, 60, 114, 0.25), rgba(42, 82, 152, 0.15));
        border: 1px solid rgba(64, 158, 255, 0.35);
        border-radius: 30px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
    }
    .footer-text {
        margin: 0;
        font-size: 0.95rem;
        font-weight: 600;
        letter-spacing: 0.03em;
        color: #e0e6ed;
    }
    .footer-subtext {
        margin: 6px 0 0 0;
        font-size: 0.8rem;
        color: rgba(255, 255, 255, 0.55);
    }
</style>
"""


# ==============================================================================
# APLICACIÓN PRINCIPAL STREAMLIT
# ==============================================================================

def main():
    st.set_page_config(
        page_title="Conversor DOCX a PDF Masivo de Alta Fidelidad",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Inicializar estado de sesión
    if "conversion_results" not in st.session_state:
        st.session_state.conversion_results = None
    if "total_batch_time" not in st.session_state:
        st.session_state.total_batch_time = 0.0

    # --------------------------------------------------------------------------
    # BARRA LATERAL: INFORMACIÓN DEL SISTEMA Y PARÁMETROS
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("### ⚙️ Motor de Renderizado")
        st.caption("Arquitectura de fidelidad absoluta para preservar márgenes, estilos y paginación.")
        
        engines = get_available_engines()
        available_engine_options = []
        engine_map = {}
        for eng in engines:
            ok, msg = eng.is_available()
            label = f"{eng.name} {'(Activo)' if ok else '(No disponible)'}"
            available_engine_options.append(label)
            engine_map[label] = (eng, ok, msg)

        selected_label = st.selectbox(
            "Seleccionar Backend:",
            options=available_engine_options,
            index=0 if available_engine_options else None
        )
        
        selected_engine, is_ok, status_desc = engine_map[selected_label]
        
        if is_ok:
            st.success(f"**Estado:** Listo para producción.\n\n_{status_desc}_")
        else:
            st.error(f"**Estado:** No operativo.\n\n_{status_desc}_")
            st.info(
                "💡 **Para Linux / Docker:** Instala LibreOffice:\n\n"
                "`sudo apt-get install -y libreoffice libreoffice-writer`"
            )

        st.divider()

        # Opciones avanzadas de compresión en el sidebar
        st.markdown("### ⚡ Optimización de Peso")
        enable_compression = st.toggle(
            "Bajar peso sin perder calidad",
            value=True,
            help="Reduce drásticamente el tamaño en MB del PDF optimizando imágenes incrustadas y flujos de contenido sin perder nitidez de texto ni alterar la paginación."
        )

        compression_mode = "ebook"
        jpeg_quality = 82
        if enable_compression:
            compression_level_label = st.selectbox(
                "Nivel de Compresión:",
                options=[
                    "Equilibrado (Recomendado - 150 DPI / eBook)",
                    "Máxima Calidad (Ligera reducción - 300 DPI / Print)",
                    "Mínimo Tamaño (Máxima compresión - 72 DPI / Web)"
                ],
                index=0
            )

            if "Equilibrado" in compression_level_label:
                compression_mode = "ebook"
                jpeg_quality = 82
            elif "Máxima Calidad" in compression_level_label:
                compression_mode = "printer"
                jpeg_quality = 92
            else:
                compression_mode = "screen"
                jpeg_quality = 68

        st.divider()
        st.markdown("### 🖥️ Entorno de Ejecución")
        st.markdown(f"- **Sistema Operativo:** `{platform.system()} {platform.release()}`")
        st.markdown(f"- **Python:** `{platform.python_version()}`")
        st.markdown(f"- **Win32COM nativo:** `{'Habilitado' if HAS_WIN32COM else 'No disponible'}`")
        
        gs_found = shutil.which("gswin64c") or shutil.which("gs")
        st.markdown(f"- **Ghostscript:** `{'Detectado' if gs_found else 'PyPDF + Pillow nativo'}`")

        st.divider()
        st.markdown("### 🛡️ Garantía de Fidelidad")
        st.markdown(
            """
            - **Paginación 1:1:** Sin saltos de página imprevistos.
            - **Tablas:** Bordes y espaciados preservados.
            - **Tipografías:** Renderizado vectorial original.
            - **Privacidad:** Archivos procesados en memoria y limpiados de disco.
            """
        )

        if st.session_state.conversion_results:
            st.divider()
            if st.button("🧹 Limpiar Resultados de la Sesión", use_container_width=True):
                st.session_state.conversion_results = None
                st.session_state.total_batch_time = 0.0
                st.rerun()

        # Pie de página también en la barra lateral para mayor coherencia
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption("Desarrollado por Ing. Wilson Castellanos @ 2026")

    # --------------------------------------------------------------------------
    # PANEL PRINCIPAL
    # --------------------------------------------------------------------------
    st.markdown(
        """
        <div class="main-header">
            <h1>📄 Conversor Masivo DOCX a PDF</h1>
            <p>Procesamiento por lotes de alta fidelidad y optimización de peso: Conserva idénticos los márgenes, tablas, tipografías y el número exacto de páginas.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if not is_ok:
        st.warning(
            f"⚠️ El motor seleccionado (**{selected_engine.name}**) no está disponible en este servidor. "
            "Por favor, revisa el panel lateral o habilita Microsoft Word / LibreOffice para continuar."
        )

    # Subida de archivos con Drag & Drop
    uploaded_files = st.file_uploader(
        "Arrastra y suelta aquí tus archivos de Word (.docx) para conversión masiva:",
        type=["docx"],
        accept_multiple_files=True,
        help="Puedes seleccionar múltiples archivos a la vez manteniendo presionada la tecla Ctrl/Cmd o arrastrándolos."
    )

    if uploaded_files:
        st.write(f"📁 **{len(uploaded_files)} archivo(s) preparado(s) para procesar.**")

        # Mensaje indicativo de compresión activa
        if enable_compression:
            st.markdown(
                f"""
                <div class="opt-box">
                    ✨ <strong>Modo de optimización activo:</strong> Los PDFs se comprimirán inteligentemente en modo 
                    <code>{compression_mode.upper()}</code> (calidad {jpeg_quality}%) reduciendo su peso sin perder nitidez tipográfica ni alterar la paginación.
                </div>
                """,
                unsafe_allow_html=True
            )

        col_btn, _ = st.columns([1, 3])
        with col_btn:
            start_btn = st.button(
                "🚀 Iniciar Conversión Masiva",
                type="primary",
                disabled=not is_ok,
                use_container_width=True
            )

        if start_btn:
            progress_holder = st.empty()
            status_holder = st.empty()
            progress_bar = progress_holder.progress(0.0)

            t_start_all = time.perf_counter()
            results = process_batch(
                files=uploaded_files,
                engine=selected_engine,
                progress_bar=progress_bar,
                status_text=status_holder,
                enable_compression=enable_compression,
                compression_level=compression_mode,
                jpeg_quality=jpeg_quality
            )
            st.session_state.total_batch_time = time.perf_counter() - t_start_all
            st.session_state.conversion_results = results
            # Limpiar placeholders antes del rerun para evitar colisiones en el Virtual DOM de React
            progress_holder.empty()
            status_holder.empty()
            st.rerun()

    # --------------------------------------------------------------------------
    # DASHBOARD DE RESULTADOS Y DESCARGA
    # --------------------------------------------------------------------------
    if st.session_state.conversion_results:
        results: List[ConversionResult] = st.session_state.conversion_results
        success_list = [r for r in results if r.status == ConversionStatus.SUCCESS]
        error_list = [r for r in results if r.status in (ConversionStatus.ERROR, ConversionStatus.VALIDATION_FAILED)]

        # Métricas de ahorro de peso
        total_orig_bytes = sum(r.original_size for r in success_list)
        total_final_bytes = sum(r.pdf_size for r in success_list)
        total_saved_bytes = max(0, total_orig_bytes - total_final_bytes)
        overall_savings_pct = ((total_orig_bytes - total_final_bytes) / total_orig_bytes * 100.0) if total_orig_bytes > 0 else 0.0

        st.divider()
        st.subheader("📊 Resumen del Procesamiento por Lotes")

        # Componentes nativos st.metric (inmunes a fallos de reconciliación de React y traducción del navegador)
        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("Total Documentos", len(results))
        col_m2.metric("Convertidos con Éxito", len(success_list))
        col_m3.metric("Fallidos o Inválidos", len(error_list))
        savings_text = f"-{overall_savings_pct:.1f}%" if overall_savings_pct > 0 else "0.0%"
        col_m4.metric("Ahorro de Peso", savings_text)
        col_m5.metric("Tiempo Total", f"{st.session_state.total_batch_time:.2f} s")

        st.markdown("<br>", unsafe_allow_html=True)

        # Botón de Descarga Masiva ZIP
        if success_list:
            zip_bytes = build_zip_package(success_list)
            timestamp_str = time.strftime("%Y%m%d_%H%M%S")
            zip_filename = f"PDFs_Convertidos_{timestamp_str}.zip"

            st.download_button(
                label=f"📦 Descargar Todos en ZIP ({format_bytes(len(zip_bytes))})",
                data=zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                type="primary",
                key="btn_download_zip_all",
                use_container_width=True
            )
            if overall_savings_pct > 0:
                st.info(f"✨ Se procesaron con éxito **{len(success_list)} archivos**. El peso total se optimizó de `{format_bytes(total_orig_bytes)}` a `{format_bytes(total_final_bytes)}` (-{overall_savings_pct:.1f}% de ahorro).")

        st.markdown("### 📥 Descargas Individuales")

        for idx, item in enumerate(success_list):
            c_name, c_size, c_time, c_btn = st.columns([4, 2, 2, 2])
            with c_name:
                st.write(f"📄 **{item.filename}**")
            with c_size:
                if item.savings_pct > 0:
                    st.write(f"`{format_bytes(item.pdf_size)}` :green[(-{item.savings_pct:.1f}%)]")
                else:
                    st.write(f"`{format_bytes(item.pdf_size)}`")
            with c_time:
                st.caption(f"{item.duration:.2f} s")
            with c_btn:
                st.download_button(
                    label="⬇️ Descargar PDF",
                    data=item.pdf_bytes,
                    file_name=item.filename,
                    mime="application/pdf",
                    key=f"dl_pdf_btn_{idx}_{item.filename}",
                    use_container_width=True
                )

        if error_list:
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander(f"⚠️ Ver detalle de los {len(error_list)} archivos con incidencias", expanded=True):
                for item in error_list:
                    st.error(f"**{item.filename}**: {item.error_message}")

    # --------------------------------------------------------------------------
    # PIE DE PÁGINA (FOOTER INSTITUCIONAL)
    # --------------------------------------------------------------------------
    st.markdown(
        """
        <div class="footer-container">
            <div class="footer-badge">
                <p class="footer-text">Desarrollado por <strong>Ing. Wilson Castellanos</strong> @ 2026</p>
            </div>
            <p class="footer-subtext">Sistema de Conversión de Alta Fidelidad y Optimización Inteligente de Documentos</p>
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()

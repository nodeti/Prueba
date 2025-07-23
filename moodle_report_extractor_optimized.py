#!/usr/bin/env python3
"""
EXTRACTOR DE REPORTES ACADÉMICOS DE MOODLE
=========================================

Este script se conecta a la API de Moodle para extraer información académica
de estudiantes de un curso específico y generar un reporte en Excel.

Funcionalidades:
- Obtiene lista de estudiantes matriculados en un curso
- Calcula progreso de actividades (completadas vs totales)
- Extrae calificaciones finales del curso
- Determina estado de aprobación (criterios: >50% actividades + nota ≥7)
- Genera reporte Excel con fecha automática
- Compatible con Google Colab (descarga automática)

Autor: Optimizado para mejor rendimiento y mantenibilidad
"""

import requests
import time
import json
import pandas as pd
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
import logging

# ===============================
# CONFIGURACIÓN PRINCIPAL
# ===============================

class MoodleConfig:
    """Configuración centralizada para la conexión a Moodle"""
    
    # URLs y credenciales
    MOODLE_URL = "https://mooc.andinavirtual.uasb.edu.ec/pruebas/webservice/rest/server.php"
    TOKEN = "ce4b68b4af3c31c3d8b434bc775"  # ⚠️ REEMPLAZAR con tu token real
    
    # Parámetros del curso
    COURSE_ID = 669
    
    # Control de rate limiting (para evitar sobrecarga del servidor)
    REQUESTS_PER_BATCH = 50    # Número de solicitudes antes de pausa
    SLEEP_SECONDS = 2          # Segundos de espera entre lotes
    
    # Criterios de aprobación
    MIN_COMPLETION_PERCENTAGE = 50  # Mínimo % de actividades completadas
    MIN_PASSING_GRADE = 7.0         # Nota mínima para aprobar
    
    # Configuración de logging
    LOG_LEVEL = logging.INFO

# ===============================
# CONFIGURACIÓN DE LOGGING
# ===============================

logging.basicConfig(
    level=MoodleConfig.LOG_LEVEL,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

# ===============================
# CLASE PRINCIPAL DEL EXTRACTOR
# ===============================

class MoodleReportExtractor:
    """
    Extractor principal para generar reportes académicos desde Moodle
    """
    
    def __init__(self, config: MoodleConfig = None):
        """
        Inicializa el extractor con la configuración proporcionada
        
        Args:
            config: Instancia de MoodleConfig (usa default si no se proporciona)
        """
        self.config = config or MoodleConfig()
        self.session = requests.Session()  # Reutilizar conexiones HTTP
        self.activity_names_cache = {}     # Cache para nombres de actividades
        
    def make_api_call(self, ws_function: str, extra_params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Realiza una llamada a la API de Moodle con manejo de errores robusto
        
        Args:
            ws_function: Nombre de la función del webservice de Moodle
            extra_params: Parámetros adicionales específicos de la función
            
        Returns:
            Respuesta JSON de la API
            
        Raises:
            Exception: Si la llamada a la API falla
        """
        # Parámetros base requeridos por Moodle
        base_params = {
            "wstoken": self.config.TOKEN,
            "wsfunction": ws_function,
            "moodlewsrestformat": "json"
        }
        
        # Combinar parámetros base con los específicos
        params = {**base_params, **extra_params}
        
        try:
            logger.debug(f"Llamando a API: {ws_function}")
            response = self.session.get(self.config.MOODLE_URL, params=params, timeout=30)
            response.raise_for_status()  # Lanza excepción si status HTTP != 200
            
            data = response.json()
            
            # Verificar si Moodle retornó un error
            if isinstance(data, dict) and "exception" in data:
                raise Exception(f"Error de Moodle: {data.get('message', 'Error desconocido')}")
                
            return data
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error de red en {ws_function}: {e}")
            raise Exception(f"Error de conexión: {e}")
        except json.JSONDecodeError as e:
            logger.error(f"Error al decodificar JSON en {ws_function}: {e}")
            raise Exception(f"Respuesta inválida del servidor: {e}")
    
    def get_activity_names_mapping(self) -> Dict[str, str]:
        """
        Obtiene el mapeo de IDs de actividades a nombres legibles
        
        Returns:
            Diccionario con mapping cmid_XXX -> nombre_actividad
        """
        if self.activity_names_cache:
            return self.activity_names_cache
            
        logger.info("🔍 Obteniendo información de actividades del curso...")
        
        try:
            course_content = self.make_api_call("core_course_get_contents", 
                                              {"courseid": self.config.COURSE_ID})
            
            activity_mapping = {}
            
            if not isinstance(course_content, list):
                raise Exception("Formato inesperado en contenido del curso")
            
            # Iterar por secciones del curso
            for section in course_content:
                section_name = section.get("name", "Sección sin nombre")
                logger.debug(f"Procesando sección: {section_name}")
                
                # Iterar por módulos/actividades en cada sección
                for module in section.get("modules", []):
                    cmid = module.get("id")
                    activity_name = module.get("name", f"Actividad_{cmid}")
                    module_type = module.get("modname", "desconocido")
                    
                    # Crear clave única para la actividad
                    key = f"cmid_{cmid}"
                    # Incluir tipo de actividad para mayor claridad
                    display_name = f"{activity_name} ({module_type})"
                    
                    activity_mapping[key] = display_name
                    logger.debug(f"Mapeada actividad: {key} -> {display_name}")
            
            self.activity_names_cache = activity_mapping
            logger.info(f"✅ Mapeadas {len(activity_mapping)} actividades")
            return activity_mapping
            
        except Exception as e:
            logger.error(f"❌ Error obteniendo actividades: {e}")
            raise Exception(f"No se pudo obtener información de actividades: {e}")
    
    def get_enrolled_students(self) -> List[Dict[str, Any]]:
        """
        Obtiene la lista de estudiantes matriculados en el curso
        
        Returns:
            Lista de diccionarios con información de estudiantes
        """
        logger.info("👥 Obteniendo lista de estudiantes matriculados...")
        
        try:
            all_users = self.make_api_call("core_enrol_get_enrolled_users", 
                                         {"courseid": self.config.COURSE_ID})
            
            # Filtrar solo estudiantes (roleid = 5 en Moodle estándar)
            students = []
            for user in all_users:
                user_roles = user.get("roles", [])
                is_student = any(role.get("roleid") == 5 for role in user_roles)
                
                if is_student:
                    students.append(user)
                    logger.debug(f"Estudiante encontrado: {user.get('fullname', 'Sin nombre')}")
            
            logger.info(f"✅ Encontrados {len(students)} estudiantes")
            return students
            
        except Exception as e:
            logger.error(f"❌ Error obteniendo estudiantes: {e}")
            raise Exception(f"No se pudo obtener lista de estudiantes: {e}")
    
    def get_student_progress(self, user_id: int, activity_mapping: Dict[str, str]) -> Tuple[Dict[str, Any], float]:
        """
        Obtiene el progreso de actividades de un estudiante específico
        
        Args:
            user_id: ID del usuario en Moodle
            activity_mapping: Mapeo de IDs a nombres de actividades
            
        Returns:
            Tupla con (progreso_por_actividad, porcentaje_completado)
        """
        try:
            progress_data = self.make_api_call("core_completion_get_activities_completion_status", {
                "courseid": self.config.COURSE_ID,
                "userid": user_id
            })
            
            activity_progress = {}
            completed_count = 0
            total_tracked_activities = 0
            
            # Procesar estado de cada actividad
            for activity in progress_data.get("statuses", []):
                completion_state = activity.get("state")
                cmid_key = f"cmid_{activity['cmid']}"
                
                # Usar nombre mapeado o clave como fallback
                activity_name = activity_mapping.get(cmid_key, cmid_key)
                
                # Convertir estado numérico a texto legible
                state_text = self._get_completion_state_text(completion_state)
                activity_progress[activity_name] = state_text
                
                # Contar para cálculo de porcentaje (solo actividades con seguimiento)
                if completion_state in (0, 1):  # 0=No completada, 1=Completada
                    total_tracked_activities += 1
                    if completion_state == 1:
                        completed_count += 1
            
            # Calcular porcentaje de completitud
            completion_percentage = 0.0
            if total_tracked_activities > 0:
                completion_percentage = round((completed_count / total_tracked_activities) * 100, 2)
            
            return activity_progress, completion_percentage
            
        except Exception as e:
            logger.warning(f"⚠️ Error obteniendo progreso para usuario {user_id}: {e}")
            return {}, 0.0
    
    def get_student_final_grade(self, user_id: int) -> Tuple[Optional[float], str]:
        """
        Obtiene la calificación final del curso para un estudiante
        
        Args:
            user_id: ID del usuario en Moodle
            
        Returns:
            Tupla con (nota_numérica, nota_formateada)
        """
        try:
            grade_data = self.make_api_call("gradereport_user_get_grade_items", {
                "courseid": self.config.COURSE_ID,
                "userid": user_id
            })
            
            # Navegar por la estructura de respuesta para encontrar la nota del curso
            user_grades = grade_data.get("usergrades", [])
            if not user_grades:
                return None, "Sin calificaciones"
            
            grade_items = user_grades[0].get("gradeitems", [])
            if not grade_items:
                return None, "Sin elementos de calificación"
            
            # El primer elemento suele ser la calificación general del curso
            course_grade = grade_items[0]
            raw_grade = course_grade.get("graderaw")
            
            if raw_grade is not None:
                formatted_grade = round(float(raw_grade), 2)
                return raw_grade, str(formatted_grade)
            else:
                return None, "No evaluado"
                
        except Exception as e:
            logger.warning(f"⚠️ Error obteniendo calificación para usuario {user_id}: {e}")
            return None, "Error al obtener nota"
    
    def determine_final_status(self, completion_percentage: float, 
                             numeric_grade: Optional[float]) -> str:
        """
        Determina el estado final del estudiante basado en criterios de aprobación
        
        Args:
            completion_percentage: Porcentaje de actividades completadas
            numeric_grade: Nota numérica final (puede ser None)
            
        Returns:
            Estado final: "Aprobado" o "Reprobado"
        """
        # Verificar criterio de actividades completadas
        meets_completion_criteria = completion_percentage >= self.config.MIN_COMPLETION_PERCENTAGE
        
        # Verificar criterio de nota mínima
        meets_grade_criteria = (numeric_grade is not None and 
                              numeric_grade >= self.config.MIN_PASSING_GRADE)
        
        # Aprobar solo si cumple ambos criterios
        if meets_completion_criteria and meets_grade_criteria:
            return "Aprobado"
        else:
            return "Reprobado"
    
    def extract_complete_report(self) -> pd.DataFrame:
        """
        Método principal que ejecuta todo el proceso de extracción
        
        Returns:
            DataFrame de pandas con el reporte completo
        """
        logger.info("🚀 Iniciando extracción completa del reporte...")
        
        # Paso 1: Obtener mapeo de actividades
        activity_mapping = self.get_activity_names_mapping()
        
        # Paso 2: Obtener lista de estudiantes
        students = self.get_enrolled_students()
        
        if not students:
            raise Exception("No se encontraron estudiantes en el curso")
        
        # Paso 3: Procesar cada estudiante
        logger.info("📊 Procesando progreso y calificaciones de estudiantes...")
        
        report_data = []
        total_students = len(students)
        
        for index, student in enumerate(students, start=1):
            user_id = student["id"]
            full_name = student.get("fullname", f"Usuario_{user_id}")
            username = student.get("username", "")
            email = student.get("email", "")
            
            logger.info(f"[{index}/{total_students}] Procesando: {full_name}")
            
            # Obtener progreso de actividades
            activity_progress, completion_percentage = self.get_student_progress(
                user_id, activity_mapping
            )
            
            # Obtener calificación final
            numeric_grade, formatted_grade = self.get_student_final_grade(user_id)
            
            # Determinar estado final
            final_status = self.determine_final_status(completion_percentage, numeric_grade)
            
            # Construir fila de datos
            student_data = {
                "ID Usuario": user_id,
                "Nombre de Usuario": username,
                "Email": email,
                "Nombre Completo": full_name,
                "% Actividades Completadas": completion_percentage,
                "Calificación Final": formatted_grade,
                "Estado Final": final_status
            }
            
            # Agregar progreso individual de actividades
            student_data.update(activity_progress)
            
            report_data.append(student_data)
            
            # Rate limiting: pausa cada cierto número de estudiantes
            if index % self.config.REQUESTS_PER_BATCH == 0 and index < total_students:
                logger.info(f"⏸️ Pausa de {self.config.SLEEP_SECONDS}s para evitar sobrecarga del servidor...")
                time.sleep(self.config.SLEEP_SECONDS)
        
        logger.info("✅ Procesamiento de estudiantes completado")
        
        # Crear DataFrame y reorganizar columnas
        df = pd.DataFrame(report_data)
        
        # Definir orden de columnas principales
        main_columns = [
            "ID Usuario", "Nombre de Usuario", "Email", "Nombre Completo",
            "% Actividades Completadas", "Calificación Final", "Estado Final"
        ]
        
        # Obtener columnas de actividades (las que no están en main_columns)
        activity_columns = [col for col in df.columns if col not in main_columns]
        activity_columns.sort()  # Ordenar alfabéticamente
        
        # Reorganizar DataFrame
        final_column_order = main_columns + activity_columns
        df = df.reindex(columns=final_column_order)
        
        logger.info(f"📈 Reporte generado: {len(df)} estudiantes, {len(df.columns)} columnas")
        return df
    
    def save_report_to_excel(self, df: pd.DataFrame, 
                           custom_filename: Optional[str] = None) -> str:
        """
        Guarda el reporte en formato Excel con nombre automático
        
        Args:
            df: DataFrame con los datos del reporte
            custom_filename: Nombre personalizado (opcional)
            
        Returns:
            Nombre del archivo generado
        """
        if custom_filename:
            filename = custom_filename
        else:
            # Generar nombre automático con fecha
            current_date = datetime.now()
            date_str = current_date.strftime("%d_%B").lower()
            
            # Normalizar caracteres especiales en español
            date_str = (date_str
                       .replace("á", "a").replace("é", "e").replace("í", "i")
                       .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
            
            filename = f"reporte_academico_curso_{self.config.COURSE_ID}_{date_str}.xlsx"
        
        # Guardar archivo
        logger.info(f"💾 Guardando reporte como: {filename}")
        
        try:
            # Configurar formato Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Reporte Académico', index=False)
                
                # Opcional: ajustar ancho de columnas
                worksheet = writer.sheets['Reporte Académico']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)  # Máximo 50 caracteres
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            logger.info(f"✅ Archivo guardado exitosamente: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"❌ Error guardando archivo: {e}")
            raise Exception(f"No se pudo guardar el archivo: {e}")
    
    @staticmethod
    def _get_completion_state_text(state: int) -> str:
        """
        Convierte estado numérico de completitud a texto legible
        
        Args:
            state: Estado numérico de Moodle
            
        Returns:
            Descripción textual del estado
        """
        state_mapping = {
            0: "No completada",
            1: "Completada",
            2: "Completada con pase",
            3: "Completada sin pase"
        }
        return state_mapping.get(state, f"Estado desconocido ({state})")

# ===============================
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# ===============================

def main():
    """
    Función principal que ejecuta todo el proceso de extracción
    """
    try:
        # Mostrar información inicial
        logger.info("=" * 60)
        logger.info("🎓 EXTRACTOR DE REPORTES ACADÉMICOS DE MOODLE")
        logger.info("=" * 60)
        logger.info(f"📚 Curso ID: {MoodleConfig.COURSE_ID}")
        logger.info(f"🌐 URL Moodle: {MoodleConfig.MOODLE_URL}")
        logger.info(f"⚙️ Configuración: {MoodleConfig.REQUESTS_PER_BATCH} requests/lote, pausa {MoodleConfig.SLEEP_SECONDS}s")
        logger.info("-" * 60)
        
        # Crear instancia del extractor
        extractor = MoodleReportExtractor()
        
        # Ejecutar extracción completa
        df_report = extractor.extract_complete_report()
        
        # Guardar en Excel
        filename = extractor.save_report_to_excel(df_report)
        
        # Mostrar resumen final
        logger.info("=" * 60)
        logger.info("📊 RESUMEN DEL REPORTE GENERADO")
        logger.info("=" * 60)
        logger.info(f"📝 Archivo: {filename}")
        logger.info(f"👥 Total estudiantes: {len(df_report)}")
        logger.info(f"📋 Total columnas: {len(df_report.columns)}")
        
        # Estadísticas de aprobación
        if "Estado Final" in df_report.columns:
            aprobados = len(df_report[df_report["Estado Final"] == "Aprobado"])
            reprobados = len(df_report[df_report["Estado Final"] == "Reprobado"])
            tasa_aprobacion = round((aprobados / len(df_report)) * 100, 1)
            
            logger.info(f"✅ Aprobados: {aprobados} ({tasa_aprobacion}%)")
            logger.info(f"❌ Reprobados: {reprobados}")
        
        logger.info("=" * 60)
        logger.info("🎉 ¡Extracción completada exitosamente!")
        
        # Intentar descarga automática (para Google Colab)
        try:
            from google.colab import files
            logger.info("📥 Iniciando descarga automática...")
            files.download(filename)
            logger.info("✅ Descarga iniciada")
        except ImportError:
            logger.info("ℹ️ Ejecutándose fuera de Google Colab - descarga manual necesaria")
        except Exception as e:
            logger.warning(f"⚠️ No se pudo iniciar descarga automática: {e}")
        
        return df_report, filename
        
    except Exception as e:
        logger.error("=" * 60)
        logger.error("💥 ERROR CRÍTICO EN LA EJECUCIÓN")
        logger.error("=" * 60)
        logger.error(f"❌ {e}")
        logger.error("=" * 60)
        raise

# ===============================
# PUNTO DE ENTRADA DEL SCRIPT
# ===============================

if __name__ == "__main__":
    # Ejecutar función principal
    report_df, output_filename = main()
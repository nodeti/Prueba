#!/usr/bin/env python3
"""
📊 GENERADOR DE REPORTES DE PROGRESO ACADÉMICO - MOODLE
================================================================
Extrae información completa del progreso de estudiantes en un curso Moodle
usando la API REST. Genera un reporte Excel con actividades, calificaciones
y estado de aprobación.

Autor: Optimizado para mejor rendimiento y mantenibilidad
Fecha: Enero 2025
"""

import requests
import time
import json
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import logging

# ═══════════════════════════════════════════════════════════════
# 🔧 CONFIGURACIÓN DEL SISTEMA
# ═══════════════════════════════════════════════════════════════

# Configuración de la API Moodle
MOODLE_CONFIG = {
    "url": "https://mooc.andinavirtual.uasb.edu.ec/pruebas/webservice/rest/server.php",
    "token": "ce4b68b4af3c31c3d8b434bc775",  # ⚠️ CAMBIAR por tu token real
    "course_id": 669,
    "format": "json"
}

# Configuración de rendimiento
PERFORMANCE_CONFIG = {
    "batch_size": 50,        # Procesar en lotes de 50 estudiantes
    "delay_seconds": 2,      # Pausa entre lotes para no sobrecargar servidor
    "timeout": 30,           # Timeout para requests HTTP
    "max_retries": 3         # Reintentos en caso de error
}

# Configuración de aprobación
APPROVAL_CRITERIA = {
    "min_completion_percentage": 50,  # Mínimo 50% actividades completadas
    "min_grade": 7.0,                # Nota mínima 7.0
    "student_role_id": 5             # ID del rol estudiante en Moodle
}

# Configurar logging para debugging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════
# 🔌 FUNCIONES DE COMUNICACIÓN CON MOODLE API
# ═══════════════════════════════════════════════════════════════

def call_moodle_api(wsfunction: str, extra_params: Dict) -> Dict:
    """
    🌐 Realiza llamadas a la API REST de Moodle con manejo de errores robusto
    
    Args:
        wsfunction: Nombre de la función del webservice de Moodle
        extra_params: Parámetros específicos de la función
    
    Returns:
        Dict con la respuesta JSON de Moodle
    
    Raises:
        Exception: Si la API retorna error o falla la comunicación
    """
    # Preparar parámetros base para todas las llamadas
    base_params = {
        "wstoken": MOODLE_CONFIG["token"],
        "wsfunction": wsfunction,
        "moodlewsrestformat": MOODLE_CONFIG["format"]
    }
    base_params.update(extra_params)
    
    # Realizar llamada con reintentos automáticos
    for attempt in range(PERFORMANCE_CONFIG["max_retries"]):
        try:
            logger.debug(f"Llamando API: {wsfunction} (intento {attempt + 1})")
            
            response = requests.get(
                MOODLE_CONFIG["url"], 
                params=base_params,
                timeout=PERFORMANCE_CONFIG["timeout"]
            )
            response.raise_for_status()  # Lanza excepción si HTTP error
            
            data = response.json()
            
            # Verificar si Moodle retornó un error
            if isinstance(data, dict) and "exception" in data:
                raise Exception(f"Error de Moodle: {data.get('message', 'Error desconocido')}")
            
            return data
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Error en intento {attempt + 1}: {e}")
            if attempt == PERFORMANCE_CONFIG["max_retries"] - 1:
                raise Exception(f"Falló comunicación con Moodle después de {PERFORMANCE_CONFIG['max_retries']} intentos: {e}")
            time.sleep(1)  # Esperar antes del siguiente intento


# ═══════════════════════════════════════════════════════════════
# 📚 FUNCIONES DE EXTRACCIÓN DE DATOS DEL CURSO
# ═══════════════════════════════════════════════════════════════

def get_course_activities_map() -> Dict[str, str]:
    """
    📋 Extrae todas las actividades del curso y crea un mapeo cmid → nombre
    
    Returns:
        Dict donde las claves son "cmid_X" y valores son nombres de actividades
    """
    logger.info("🔍 Obteniendo actividades del curso...")
    
    course_content = call_moodle_api("core_course_get_contents", {
        "courseid": MOODLE_CONFIG["course_id"]
    })
    
    activities_map = {}
    total_activities = 0
    
    if not isinstance(course_content, list):
        raise Exception("❌ Formato inesperado en contenido del curso")
    
    # Recorrer todas las secciones del curso
    for section in course_content:
        section_name = section.get("name", "Sección sin nombre")
        logger.debug(f"Procesando sección: {section_name}")
        
        # Extraer módulos/actividades de cada sección
        for module in section.get("modules", []):
            cmid = module.get("id")
            activity_name = module.get("name", f"Actividad {cmid}")
            
            # Crear clave única para el mapeo
            cmid_key = f"cmid_{cmid}"
            activities_map[cmid_key] = activity_name
            total_activities += 1
    
    logger.info(f"✅ Se encontraron {total_activities} actividades en el curso")
    return activities_map


def get_enrolled_students() -> List[Dict]:
    """
    👥 Obtiene lista de usuarios matriculados y filtra solo estudiantes
    
    Returns:
        Lista de diccionarios con información de estudiantes
    """
    logger.info("👥 Obteniendo estudiantes matriculados...")
    
    enrolled_users = call_moodle_api("core_enrol_get_enrolled_users", {
        "courseid": MOODLE_CONFIG["course_id"]
    })
    
    # Filtrar solo usuarios con rol de estudiante (roleid = 5)
    students = []
    for user in enrolled_users:
        user_roles = user.get("roles", [])
        is_student = any(role["roleid"] == APPROVAL_CRITERIA["student_role_id"] for role in user_roles)
        
        if is_student:
            students.append(user)
    
    logger.info(f"✅ Se encontraron {len(students)} estudiantes")
    return students


# ═══════════════════════════════════════════════════════════════
# 📊 FUNCIONES DE ANÁLISIS DE PROGRESO Y CALIFICACIONES
# ═══════════════════════════════════════════════════════════════

def get_student_completion_data(student_id: int, activities_map: Dict[str, str]) -> Tuple[Dict, float]:
    """
    📈 Extrae datos de completitud de actividades para un estudiante específico
    
    Args:
        student_id: ID único del estudiante en Moodle
        activities_map: Mapeo de cmid a nombres de actividades
    
    Returns:
        Tupla con (datos_completitud, porcentaje_completado)
    """
    completion_data = call_moodle_api("core_completion_get_activities_completion_status", {
        "courseid": MOODLE_CONFIG["course_id"],
        "userid": student_id
    })
    
    # Inicializar contadores y datos de actividades
    completion_row = {}
    completed_count = 0
    total_tracked_activities = 0
    
    # Procesar estado de cada actividad
    for activity_status in completion_data.get("statuses", []):
        completion_state = activity_status.get("state")
        cmid_key = f"cmid_{activity_status['cmid']}"
        
        # Usar nombre legible de actividad como columna
        column_name = activities_map.get(cmid_key, cmid_key)
        completion_row[column_name] = completion_state
        
        # Contar solo actividades con seguimiento de completitud
        if completion_state in (0, 1):  # 0=incompleta, 1=completa
            total_tracked_activities += 1
            if completion_state == 1:
                completed_count += 1
    
    # Calcular porcentaje de completitud
    completion_percentage = (
        round((completed_count / total_tracked_activities) * 100, 2) 
        if total_tracked_activities > 0 else 0
    )
    
    return completion_row, completion_percentage


def get_student_final_grade(student_id: int) -> Tuple[Optional[float], str]:
    """
    🎯 Obtiene la calificación final del curso para un estudiante
    
    Args:
        student_id: ID único del estudiante en Moodle
    
    Returns:
        Tupla con (nota_numérica, nota_formateada)
    """
    try:
        grade_report = call_moodle_api("gradereport_user_get_grade_items", {
            "courseid": MOODLE_CONFIG["course_id"],
            "userid": student_id
        })
        
        # Navegar estructura de calificaciones: usergrades[0].gradeitems[0]
        user_grades = grade_report.get("usergrades", [])
        if not user_grades:
            return None, "N/A"
        
        grade_items = user_grades[0].get("gradeitems", [])
        if not grade_items:
            return None, "N/A"
        
        # Obtener calificación bruta del curso
        course_grade_item = grade_items[0]
        raw_grade = course_grade_item.get("graderaw")
        
        if raw_grade is not None:
            return raw_grade, str(round(raw_grade, 2))
        else:
            return None, "N/A"
            
    except Exception as e:
        logger.warning(f"Error obteniendo calificación para estudiante {student_id}: {e}")
        return None, "Error"


def determine_final_status(completion_percentage: float, numeric_grade: Optional[float]) -> str:
    """
    🏆 Determina el estado final del estudiante basado en criterios de aprobación
    
    Args:
        completion_percentage: Porcentaje de actividades completadas
        numeric_grade: Calificación numérica final (puede ser None)
    
    Returns:
        "Aprobado" o "Reprobado"
    """
    meets_completion = completion_percentage >= APPROVAL_CRITERIA["min_completion_percentage"]
    meets_grade = numeric_grade is not None and numeric_grade >= APPROVAL_CRITERIA["min_grade"]
    
    return "Aprobado" if (meets_completion and meets_grade) else "Reprobado"


# ═══════════════════════════════════════════════════════════════
# 🎯 FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# ═══════════════════════════════════════════════════════════════

def process_all_students(students: List[Dict], activities_map: Dict[str, str]) -> List[Dict]:
    """
    🔄 Procesa el progreso académico de todos los estudiantes
    
    Args:
        students: Lista de estudiantes a procesar
        activities_map: Mapeo de actividades del curso
    
    Returns:
        Lista con datos completos de progreso de cada estudiante
    """
    logger.info(f"🔄 Iniciando procesamiento de {len(students)} estudiantes...")
    
    progress_data = []
    batch_size = PERFORMANCE_CONFIG["batch_size"]
    delay = PERFORMANCE_CONFIG["delay_seconds"]
    
    for index, student in enumerate(students, start=1):
        # Extraer información básica del estudiante
        student_info = {
            "userid": student["id"],
            "username": student.get("username", ""),
            "email": student.get("email", ""),
            "fullname": student.get("fullname", f"Usuario {student['id']}")
        }
        
        try:
            # Obtener datos de completitud de actividades
            completion_data, completion_percentage = get_student_completion_data(
                student["id"], activities_map
            )
            student_info.update(completion_data)
            student_info["% actividades completadas"] = completion_percentage
            
            # Obtener calificación final
            numeric_grade, formatted_grade = get_student_final_grade(student["id"])
            student_info["calificación final"] = formatted_grade
            
            # Determinar estado final (aprobado/reprobado)
            final_status = determine_final_status(completion_percentage, numeric_grade)
            student_info["estado final"] = final_status
            
            progress_data.append(student_info)
            
            # Log de progreso
            logger.info(f"[{index}/{len(students)}] ✅ {student_info['fullname']} - "
                       f"{completion_percentage}% - {formatted_grade} - {final_status}")
            
        except Exception as e:
            # Manejo de errores por estudiante individual
            logger.error(f"❌ Error procesando {student_info['fullname']}: {e}")
            student_info.update({
                "% actividades completadas": "Error",
                "calificación final": "Error",
                "estado final": "Error"
            })
            progress_data.append(student_info)
        
        # Pausa periódica para no sobrecargar el servidor
        if index % batch_size == 0 and index < len(students):
            logger.info(f"⏸️ Pausa de {delay} segundos para optimizar rendimiento...")
            time.sleep(delay)
    
    logger.info("✅ Procesamiento completado")
    return progress_data


# ═══════════════════════════════════════════════════════════════
# 📁 FUNCIONES DE EXPORTACIÓN Y ALMACENAMIENTO
# ═══════════════════════════════════════════════════════════════

def generate_filename() -> str:
    """
    📅 Genera nombre de archivo con fecha actual en español
    
    Returns:
        Nombre de archivo formateado: "progreso_estudiantes_curso1_DD_mes.xlsx"
    """
    # Mapeo de meses en inglés a español
    month_translation = {
        "january": "enero", "february": "febrero", "march": "marzo",
        "april": "abril", "may": "mayo", "june": "junio",
        "july": "julio", "august": "agosto", "september": "septiembre",
        "october": "octubre", "november": "noviembre", "december": "diciembre"
    }
    
    # Obtener fecha actual y formatear
    current_date = datetime.now()
    day = current_date.strftime("%d")
    month_en = current_date.strftime("%B").lower()
    month_es = month_translation.get(month_en, month_en)
    
    return f"progreso_estudiantes_curso1_{day}_{month_es}.xlsx"


def export_to_excel(progress_data: List[Dict]) -> str:
    """
    📊 Exporta los datos de progreso a un archivo Excel
    
    Args:
        progress_data: Lista con datos de progreso de todos los estudiantes
    
    Returns:
        Nombre del archivo Excel generado
    """
    logger.info("📊 Generando reporte Excel...")
    
    # Crear DataFrame con todos los datos
    df = pd.DataFrame(progress_data)
    
    # Generar nombre de archivo con fecha
    filename = generate_filename()
    
    # Guardar archivo Excel
    df.to_excel(filename, index=False, engine='openpyxl')
    
    # Estadísticas del reporte
    total_students = len(df)
    approved_students = len(df[df["estado final"] == "Aprobado"])
    approval_rate = (approved_students / total_students * 100) if total_students > 0 else 0
    
    logger.info(f"✅ Reporte generado: {filename}")
    logger.info(f"📈 Estadísticas: {total_students} estudiantes, "
               f"{approved_students} aprobados ({approval_rate:.1f}%)")
    
    return filename


# ═══════════════════════════════════════════════════════════════
# 🚀 FUNCIÓN PRINCIPAL - PUNTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════

def main():
    """
    🎯 Función principal que orquesta todo el proceso de generación de reportes
    """
    try:
        print("🚀 INICIANDO GENERADOR DE REPORTES DE PROGRESO ACADÉMICO")
        print("=" * 60)
        
        # Paso 1: Obtener mapeo de actividades del curso
        activities_map = get_course_activities_map()
        
        # Paso 2: Obtener lista de estudiantes matriculados
        students = get_enrolled_students()
        
        if not students:
            raise Exception("❌ No se encontraron estudiantes en el curso")
        
        # Paso 3: Procesar progreso de todos los estudiantes
        progress_data = process_all_students(students, activities_map)
        
        # Paso 4: Exportar datos a Excel
        filename = export_to_excel(progress_data)
        
        print("=" * 60)
        print(f"🎉 PROCESO COMPLETADO EXITOSAMENTE")
        print(f"📁 Archivo generado: {filename}")
        print(f"📊 Datos de {len(progress_data)} estudiantes procesados")
        
        # En Google Colab, descargar automáticamente
        try:
            from google.colab import files
            files.download(filename)
            print("⬇️ Descarga iniciada automáticamente")
        except ImportError:
            print("💡 Para descargar en Colab, ejecuta: files.download('{filename}')")
    
    except Exception as e:
        logger.error(f"💥 Error crítico en el proceso: {e}")
        print(f"❌ Error: {e}")
        raise


# ═══════════════════════════════════════════════════════════════
# 🏁 EJECUCIÓN DEL SCRIPT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
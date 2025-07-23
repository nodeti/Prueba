# 🎓 Extractor de Reportes Académicos de Moodle

Un script en Python optimizado para extraer reportes académicos completos desde la API de Moodle, incluyendo progreso de actividades, calificaciones y estadísticas de aprobación.

## 📋 Características

- ✅ **Extracción completa**: Estudiantes, actividades, progreso y calificaciones
- 📊 **Cálculo automático**: Porcentajes de completitud y estado de aprobación
- 📁 **Exportación Excel**: Reporte organizado con formato profesional
- 🔄 **Rate limiting**: Control inteligente para evitar sobrecarga del servidor
- 📝 **Logging detallado**: Seguimiento completo del proceso de extracción
- 🛡️ **Manejo de errores**: Recuperación robusta ante fallos de red o API
- 🌐 **Compatible Google Colab**: Descarga automática de archivos

## 🚀 Instalación Rápida

### Opción 1: Google Colab (Recomendado)
```python
# Las librerías ya están instaladas en Colab
# Solo ejecuta el script directamente
```

### Opción 2: Entorno Local
```bash
# Clonar o descargar el script
pip install -r requirements.txt
python moodle_report_extractor_optimized.py
```

## ⚙️ Configuración

### 1. Configurar Credenciales
Edita la clase `MoodleConfig` en el script:

```python
class MoodleConfig:
    # URLs y credenciales
    MOODLE_URL = "https://tu-moodle.edu/webservice/rest/server.php"
    TOKEN = "tu_token_aqui"  # ⚠️ REEMPLAZAR con tu token real
    
    # Parámetros del curso
    COURSE_ID = 669  # ID del curso a analizar
```

### 2. Obtener Token de Moodle
1. Accede a tu Moodle como administrador
2. Ve a **Administración del sitio** → **Plugins** → **Servicios web**
3. Habilita los servicios web
4. Crea un token para tu usuario
5. Asegúrate de tener permisos para estas funciones:
   - `core_course_get_contents`
   - `core_enrol_get_enrolled_users`
   - `core_completion_get_activities_completion_status`
   - `gradereport_user_get_grade_items`

### 3. Criterios de Aprobación
Personaliza los criterios según tu institución:

```python
# Criterios de aprobación
MIN_COMPLETION_PERCENTAGE = 50  # Mínimo % de actividades completadas
MIN_PASSING_GRADE = 7.0         # Nota mínima para aprobar
```

## 📖 Uso

### Ejecución Básica
```python
# Ejecutar con configuración por defecto
python moodle_report_extractor_optimized.py
```

### Uso Programático
```python
from moodle_report_extractor_optimized import MoodleReportExtractor, MoodleConfig

# Configuración personalizada
config = MoodleConfig()
config.COURSE_ID = 123
config.TOKEN = "mi_token"

# Crear extractor
extractor = MoodleReportExtractor(config)

# Generar reporte
df_report = extractor.extract_complete_report()

# Guardar con nombre personalizado
filename = extractor.save_report_to_excel(df_report, "mi_reporte_custom.xlsx")
```

## 📊 Estructura del Reporte

El archivo Excel generado contiene las siguientes columnas:

### Información Básica del Estudiante
- **ID Usuario**: Identificador único en Moodle
- **Nombre de Usuario**: Username de acceso
- **Email**: Correo electrónico
- **Nombre Completo**: Nombre y apellidos

### Métricas de Rendimiento
- **% Actividades Completadas**: Porcentaje de actividades finalizadas
- **Calificación Final**: Nota final del curso
- **Estado Final**: "Aprobado" o "Reprobado" según criterios

### Progreso Detallado por Actividad
- Una columna por cada actividad del curso
- Estados posibles:
  - "No completada"
  - "Completada" 
  - "Completada con pase"
  - "Completada sin pase"

## 🎯 Ejemplo de Salida

```
============================================================
🎓 EXTRACTOR DE REPORTES ACADÉMICOS DE MOODLE
============================================================
📚 Curso ID: 669
🌐 URL Moodle: https://mooc.andinavirtual.uasb.edu.ec/...
⚙️ Configuración: 50 requests/lote, pausa 2s
------------------------------------------------------------
🔍 Obteniendo información de actividades del curso...
✅ Mapeadas 15 actividades
👥 Obteniendo lista de estudiantes matriculados...
✅ Encontrados 45 estudiantes
📊 Procesando progreso y calificaciones de estudiantes...
[1/45] Procesando: Juan Pérez
[2/45] Procesando: María García
...
✅ Procesamiento de estudiantes completado
💾 Guardando reporte como: reporte_academico_curso_669_15_enero.xlsx
============================================================
📊 RESUMEN DEL REPORTE GENERADO
============================================================
📝 Archivo: reporte_academico_curso_669_15_enero.xlsx
👥 Total estudiantes: 45
📋 Total columnas: 22
✅ Aprobados: 32 (71.1%)
❌ Reprobados: 13
============================================================
🎉 ¡Extracción completada exitosamente!
```

## 🔧 Personalización Avanzada

### Control de Rate Limiting
```python
# Ajustar velocidad de procesamiento
REQUESTS_PER_BATCH = 25    # Reducir para servidores lentos
SLEEP_SECONDS = 5          # Aumentar pausa entre lotes
```

### Nivel de Logging
```python
import logging

# Cambiar nivel de detalle
LOG_LEVEL = logging.DEBUG  # Más detallado
LOG_LEVEL = logging.WARNING  # Solo advertencias y errores
```

### Criterios de Aprobación Personalizados
```python
def determine_final_status(self, completion_percentage, numeric_grade):
    # Ejemplo: Criterio más estricto
    if completion_percentage >= 80 and numeric_grade >= 8.0:
        return "Excelente"
    elif completion_percentage >= 60 and numeric_grade >= 7.0:
        return "Aprobado"
    elif completion_percentage >= 40 and numeric_grade >= 6.0:
        return "Suficiente"
    else:
        return "Reprobado"
```

## 🐛 Solución de Problemas

### Error: "Token inválido"
- Verifica que el token esté correctamente copiado
- Asegúrate de que el token tenga permisos para servicios web
- Confirma que los servicios web estén habilitados en Moodle

### Error: "Función no disponible"
```python
# Verificar funciones habilitadas en Moodle
# Administración → Plugins → Servicios web → Funciones externas
```

### Error: "Timeout de conexión"
```python
# Aumentar timeout en make_api_call
response = self.session.get(url, params=params, timeout=60)
```

### Reporte vacío o incompleto
- Verifica que el `COURSE_ID` sea correcto
- Confirma que hay estudiantes matriculados con rol de estudiante (roleid=5)
- Revisa que las actividades tengan seguimiento de completitud habilitado

### Performance lenta
```python
# Reducir batch size para servidores lentos
REQUESTS_PER_BATCH = 10
SLEEP_SECONDS = 3
```

## 📄 Licencia

Este script es de uso libre para fines educativos. Asegúrate de cumplir con las políticas de uso de datos de tu institución.

## 🤝 Contribuciones

Las mejoras y sugerencias son bienvenidas. Considera:
- Manejo de más tipos de actividades
- Soporte para múltiples cursos
- Exportación a otros formatos (CSV, PDF)
- Interfaz gráfica de usuario

## ⚠️ Consideraciones de Seguridad

- **Nunca compartas tu token** en repositorios públicos
- Usa variables de entorno para credenciales en producción
- Respeta los límites de rate limiting del servidor
- Cumple con las políticas de privacidad de datos estudiantiles
# 📊 Sistema Automatizado de Boletines Académicos

## 📝 Descripción del Proyecto
Herramienta de automatización desarrollada en **Google Apps Script (JavaScript)** diseñada para optimizar el flujo de trabajo en la gestión académica. El script procesa bases de datos de calificaciones en Google Sheets, inyecta la información en plantillas dinámicas de Google Docs para generar boletines individuales en formato PDF, y automatiza su distribución masiva vía correo electrónico.

Este desarrollo elimina el margen de error humano y reduce el tiempo de procesamiento manual de horas a minutos.

## 🛠️ Tecnologías Utilizadas
- **Lenguaje:** Google Apps Script (JavaScript puro adaptado al motor V8)
- **Integraciones de APIs (Google Workspace):** 
  - SpreadsheetApp (Google Sheets)
  - DocumentApp (Google Docs)
  - DriveApp (Manipulación y exportación a PDF)
  - GmailApp (Distribución automatizada)

## ⚙️ Arquitectura y Lógica del Script
1. **Lectura masiva de datos:** Extracción iterativa de registros de estudiantes y variables de calificación desde arreglos bidimensionales.
2. **Generación dinámica de documentos:** Clonación de plantillas maestras y reemplazo de etiquetas de variables (ej. `{{NOMBRE_ESTUDIANTE}}`) en tiempo de ejecución.
3. **Inmutabilidad de datos:** Exportación automática del documento generado a un Blob de tipo PDF para evitar alteraciones posteriores en las notas.
4. **Distribución asíncrona:** Envío iterativo de correos electrónicos personalizados con el boletín adjunto correspondiente a cada destinatario.

## 🚀 Instalación y Despliegue
*(Nota: El código está diseñado para ejecutarse dentro del entorno de Google Workspace).*

1. Crear un nuevo proyecto en la extensión de Apps Script desde una hoja de cálculo de Google.
2. Copiar el contenido del archivo `main.gs` en el editor.
3. Configurar las variables globales en el script:
   - `TEMPLATE_ID`: El ID del documento de Google Docs usado como plantilla.
   - `FOLDER_ID`: El ID de la carpeta de Drive donde se guardarán los PDFs.
4. Ejecutar la función principal y otorgar los permisos de OAuth requeridos por Google para acceder a Drive y Gmail.

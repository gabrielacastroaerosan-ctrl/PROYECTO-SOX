# Centro Operativo SOX de Tarifas

Versión mejorada de la Web App de Google Apps Script. Conserva el procesamiento actual y lleva el análisis, el seguimiento y las evidencias a una sola interfaz.

## Qué cambia

- Tablero ejecutivo dentro de la app con los dos hallazgos definidos por la jefatura: aprobadores no autorizados y autoaprobaciones.
- Sábana consolidada filtrable por período, tipo, zona, hallazgo y texto libre.
- Vista detallada de cada caso, evidencia y estado de revisión.
- Descarga del panel completo como archivo `.xlsx` para Excel.
- Historial de ejecuciones y estado del recordatorio quincenal.
- Portal de usuario con resumen, filtros y captura de nota o evidencia.
- Se conserva la actualización del Google Sheet maestro y la organización de archivos en Drive.

## Instalación

1. Abre el proyecto actual de Apps Script.
2. Reemplaza el contenido de `Code.gs`, `index.html` y `portal.html` con estos archivos.
3. En **Configuración del proyecto**, activa **Mostrar el archivo de manifiesto** y reemplaza `appsscript.json`.
4. Verifica en `CONFIG` los IDs del Sheet y la carpeta, reemplaza `ALEJANDRA_CORREO_AQUI@latam.com` y llena `FOCALS_EMAILS` con los destinatarios de SAM, NAM y EUR.
5. En **Servicios**, confirma que Google Drive API v3 esté habilitada.
6. Ejecuta `obtenerDashboardSOX` una vez desde el editor para autorizar los permisos.
7. Implementa una nueva versión de la aplicación web con acceso restringido al dominio.

## Operación esperada

1. Actualizar la matriz de aprobadores hasta el último mes cerrado.
2. Procesar los archivos del trimestre desde **Procesar archivos**.
3. Revisar las dos tablas críticas en **Resumen ejecutivo**.
4. Filtrar o inspeccionar registros en **Hallazgos**.
5. Enviar invitaciones desde **Comunicaciones**.
6. Aprobar o devolver la evidencia desde el detalle del caso.
7. Descargar el libro completo desde **Descargar Excel** cuando se necesite una copia fuera de Google Sheets.

## Nota de compatibilidad

Los registros históricos que no tengan la nueva columna `Periodo` aparecerán como **Período actual**. Los lotes procesados con esta versión guardan el trimestre explícitamente.

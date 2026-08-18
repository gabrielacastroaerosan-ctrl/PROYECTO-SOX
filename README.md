# Centro Operativo SOX de Tarifas

Versión mejorada de la Web App de Google Apps Script. Conserva el procesamiento actual y lleva el análisis, el seguimiento y las evidencias a una sola interfaz.

## Qué cambia

- Tablero ejecutivo dentro de la app con los dos hallazgos definidos por la jefatura: aprobadores no autorizados y autoaprobaciones.
- Sábana consolidada filtrable por período, tipo, zona, hallazgo y texto libre.
- Vista detallada de cada caso, evidencia y estado de revisión.
- Descarga del panel completo como archivo `.xlsx` para Excel.
- Historial de ejecuciones y estado del recordatorio quincenal.
- Un único correo general trimestral con la sábana consolidada y acceso al portal, donde cada usuario ve solo sus registros.
- Recordatorio general mensual a los focals, programado para el día 15 y referido al último mes cerrado.
- Portal de usuario con resumen, filtros y captura de nota o evidencia.
- Flujo guiado de cinco pasos que indica la siguiente acción del período.
- Explorador de solo lectura para consultar las carpetas y archivos del repositorio SOX en Drive desde la app.
- Se conserva la actualización del Google Sheet maestro y la organización de archivos en Drive.

## Instalación

1. Abre el proyecto actual de Apps Script.
2. Reemplaza el contenido de `Code.gs`, `index.html` y `portal.html` con estos archivos.
3. En **Configuración del proyecto**, activa **Mostrar el archivo de manifiesto** y reemplaza `appsscript.json`.
4. Verifica en `CONFIG` los IDs del Sheet y la carpeta, reemplaza `ALEJANDRA_CORREO_AQUI@latam.com`, llena `FOCALS_EMAILS` con los focals de SAM, NAM y EUR y configura `DESTINATARIOS_REPORTE` con las listas o grupos que recibirán el correo trimestral.
5. En **Servicios**, confirma que Google Drive API v3 esté habilitada.
6. Ejecuta `obtenerDashboardSOX` una vez desde el editor para autorizar los permisos.
7. Implementa una nueva versión de la aplicación web con acceso restringido al dominio.

## Operación esperada

1. Actualizar la matriz de aprobadores hasta el último mes cerrado.
2. Procesar los archivos del trimestre desde **Procesar archivos**.
3. Revisar las dos tablas críticas en **Resumen ejecutivo**.
4. Filtrar o inspeccionar registros en **Hallazgos**.
5. Consultar los archivos procesados en **Repositorio Drive**.
6. Previsualizar y enviar una sola vez el **correo general trimestral** desde **Comunicaciones**.
7. Aprobar o devolver la evidencia desde el detalle del caso.
8. Descargar el libro completo desde **Descargar Excel** cuando se necesite una copia fuera de Google Sheets.

## Nota de compatibilidad

Los registros históricos que no tengan la nueva columna `Periodo` aparecerán como **Período actual**. Los lotes procesados con esta versión guardan el trimestre explícitamente.

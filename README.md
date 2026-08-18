# Centro Operativo SOX de Tarifas

Versión mejorada de la Web App de Google Apps Script. Conserva el procesamiento actual y lleva el análisis, el seguimiento y las evidencias a una sola interfaz.

## Qué cambia

- Tablero ejecutivo dentro de la app con los dos hallazgos definidos por la jefatura: aprobadores no autorizados y autoaprobaciones.
- Sábana consolidada filtrable por período, tipo, zona, hallazgo y texto libre.
- Vista detallada de cada caso, evidencia y estado de revisión.
- Reporte trimestral separado del panel maestro, con solo dos hojas: `Resumen ejecutivo` y `Detalle consolidado`.
- Descarga de ese reporte como `.xlsx`; no se crean pestañas por persona.
- Historial de ejecuciones y estado del recordatorio quincenal.
- Un único correo general trimestral con la sábana consolidada y acceso al portal, donde cada usuario ve solo sus registros.
- Recordatorio general mensual a los focals, programado para el día 15 y referido al último mes cerrado.
- Portal de usuario con resumen, filtros y captura de nota o evidencia.
- Flujo guiado de cinco pasos que indica la siguiente acción del período.
- Indicador de conexión con animación y segundos transcurridos mientras se consultan Sheets y Drive.
- Carga inicial liviana: los KPIs se calculan sobre todo el trimestre, pero la lista descarga solo una primera página; búsquedas y filtros consultan el período completo en el servidor con una pausa breve para evitar llamadas repetidas.
- Explorador de solo lectura para consultar las carpetas y archivos del repositorio SOX en Drive desde la app.
- Se conserva la actualización del Google Sheet maestro y la organización de archivos en Drive.
- El botón **Cómo se calcula** explica en la app el alcance y la fórmula de cada estadística del trimestre seleccionado.
- La vista previa segura del correo muestra período, destinatarios y cifras exactas antes de enviar; abrirla nunca envía mensajes.
- El correo trimestral usa una plantilla corporativa con asunto de acción requerida, resumen ejecutivo, instrucciones, indicadores, detalle priorizado y accesos al portal, al reporte y al repositorio.
- La vista previa puede abrirse en una ventana completa y el envío queda visualmente bloqueado cuando no existen destinatarios configurados.
- El tablero distingue explícitamente entre **registros** y **personas únicas**. Las tablas agrupadas muestran su suma de comprobación y cada persona abre sus registros exactos en la revisión detallada.

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
2. Seleccionar explícitamente el trimestre en **Procesar período** y cargar el paquete completo; el nombre del archivo ya no decide la carpeta.
3. Revisar las dos tablas críticas en **Inicio y flujo**.
4. Filtrar o inspeccionar registros en **Revisar resultados**.
5. Consultar los archivos procesados en **Repositorio Drive**.
6. Generar o actualizar el **reporte para envío** desde **Comunicaciones** y revisarlo en Google Sheets o Excel.
7. Previsualizar y enviar una sola vez el **correo general trimestral**; incluye el enlace al reporte, el portal individual y el `.xlsx` adjunto cuando su tamaño lo permite.
8. Aprobar o devolver la evidencia desde el detalle del caso.

## Separación de entregables

- **Soporte técnico interno:** cada archivo procesado conserva su propia hoja `Análisis`; el panel maestro mantiene datos incompletos y trazabilidad para investigación.
- **Entregable para jefatura y equipos:** un solo libro trimestral, filtrable, sin hojas individuales. Los casos sin un correo de aprobador válido se excluyen de la distribución y permanecen en el análisis interno.

## Regla de almacenamiento y reproceso

- La carpeta se determina por el trimestre seleccionado (`Q1 AAAA` a `Q4 AAAA`). Si ya existe, se reutiliza; de lo contrario, la app la crea dentro del repositorio SOX.
- El selector muestra siempre los cuatro trimestres de cada año. Los períodos futuros permanecen visibles pero deshabilitados hasta que comiencen.
- Al reprocesar, se reemplaza el consolidado de ese trimestre y se conservan los demás períodos históricos.
- Dentro de la carpeta trimestral, un archivo nuevo con el mismo nombre reemplaza la versión anterior, que se envía a la papelera. Los archivos con nombres diferentes se conservan.
- Por esta razón debe cargarse el paquete completo del trimestre en una misma ejecución.
- La app no reconstruye automáticamente períodos anteriores: si se empieza a operar en `Q3`, el historial inicia en `Q3` salvo que se carguen manualmente los paquetes completos de `Q1` y `Q2`.

## Reinicio para pruebas

En **Configuración → Zona de pruebas** los administradores pueden reiniciar los datos operativos escribiendo `REINICIAR`. La acción limpia resultados, evidencias e historiales y envía las carpetas trimestrales a la papelera. No elimina la carpeta raíz ni modifica `Permisos`, `diccionario postas`, `Justificados` o la configuración del proyecto.

## Nota de compatibilidad

Los registros históricos que no tengan la nueva columna `Periodo` aparecerán como **Período actual**. Los lotes procesados con esta versión guardan el trimestre explícitamente.

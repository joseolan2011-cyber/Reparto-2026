# Arquitectura de cierres de Reparto 2026

Fecha de documentación: 2026-10-08

## Regla principal

El botón **Ver detalle** de AppSheet NO usa el Web App de Atlas del repositorio Reparto-2026.

El reporte HTML **Cierre detallado** pertenece a un proyecto Apps Script standalone separado llamado:

**CIERRE_REPARTO_APPSHEET**

Antes de modificar un reporte de cierre, confirmar primero que se está trabajando en ese proyecto.

## Proyecto 1: Reparto-2026 / Atlas

Repositorio GitHub:
- joseolan2011-cyber/Reparto-2026

Responsabilidades:
- Atlas / Web App principal.
- MonitorPedidos.
- MonitorMake.
- SYNC_CARGAS_OPSU.
- CIERRES_REPARTO de cálculo y escritura de cierres.
- Otros módulos operativos del proyecto Reparto 2026.

Archivos importantes:
- WEB_App.gs.js
- CIERRES_REPARTO.js
- Index.html

**NO usar este Web App para modificar el HTML del botón "Ver detalle" de AppSheet.**

## Proyecto 2: CIERRE_REPARTO_APPSHEET

Tipo:
- Google Apps Script standalone.

Responsabilidad:
- Cierre llamado desde AppSheet.
- Reporte HTML detallado del cierre.
- Ticket de cierre.

Archivos visibles confirmados:
- Código.gs
- DETALLE_CIERRE_WEB.gs

El archivo a revisar primero cuando se solicite cambiar el diseño o contenido del reporte web es:

**DETALLE_CIERRE_WEB.gs**

La acción de AppSheet usa un Web App de este proyecto standalone y envía:

?cierre=[id_cierre]

## Procedimiento obligatorio antes de cambios

1. Identificar qué pantalla o botón está generando el resultado.
2. Revisar la URL real que abre AppSheet.
3. Confirmar en Apps Script que el deployment pertenece a **CIERRE_REPARTO_APPSHEET**.
4. Respaldar los archivos del proyecto que se van a modificar.
5. No modificar WEB_App.gs.js de Atlas para cambios del reporte detallado.
6. No tocar Telegram, MonitorPedidos, MonitorMake ni SYNC_CARGAS_OPSU para cambios visuales del cierre.
7. Hacer primero una marca visible de versión en el reporte modificado.
8. Probar con un solo id_cierre.
9. Sólo después retirar la marca de prueba y publicar la versión final.

## Recuperación

Respaldo de GitHub anterior al incidente:

**backup-cierres-reparto-2026-10-08**

Respaldo del Google Sheet REPARTO 2026:

**RESPALDO REPARTO 2026 - antes ticket $20 y subtotales cargas - 2026-10-08**

El 2026-10-08 se restauraron desde el respaldo:
- CIERRES_REPARTO.js
- WEB_App.gs.js
- .github/workflows/deploy-appscript.yml

También se eliminaron de la hoja CIERRES_REPARTO las columnas temporales:
- Ticket_20
- Reporte_Cargas

## Regla de seguridad

No hacer cambios de código fuera del componente identificado para la solicitud. Para un cambio visual o de contenido del reporte de cierre, trabajar únicamente en **CIERRE_REPARTO_APPSHEET**, salvo que exista una razón técnica comprobada y autorizada para tocar otro proyecto.

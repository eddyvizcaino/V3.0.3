# V3.0.3 — Mejoras de Asistencia
- Búsqueda por MEM, nombre, apellido, teléfono y filtro por ministerio (se conserva).
- Contadores de miembros, visitantes y total (se conserva).
- Registro inmediato en base de datos y protección contra duplicados por culto (se conserva).
- Corrección de asistencia desde PWA, auditada, solo durante el culto para operador.
- Panel protegido `/attendance/seguimiento`: ausencias, última asistencia, porcentajes, ministerios y alerta por dos viernes consecutivos.
- Exportación CSV compatible con Excel.
- Se mantienen los cultos de viernes y domingo; se añaden claves de Culto Especial, Vigilia, Santa Cena y Conferencia para selección manual de administrador, sin creación ni edición de cultos.
- El rol Registro de Asistencia no puede forzar fechas ni cultos distintos del actual por solicitudes manipuladas.

## Límites importantes
- No se puede inferir que un culto ocurrió si nadie registró asistencia o visitantes. Solo se calculan ausencias respecto a cultos que tienen registros. Conviene incorporar más adelante un cierre explícito de culto.
- No se han implementado en esta entrega impresión automática de tickets, seguimiento de contactos con notas, PDF ni identificación de visitantes recurrentes. Requieren diseño y pruebas adicionales.
- Verificar en Render con copia de seguridad previa; pruebas locales no equivalen a validación de PostgreSQL de producción.

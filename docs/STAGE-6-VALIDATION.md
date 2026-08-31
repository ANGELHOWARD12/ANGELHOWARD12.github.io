# Etapa 6 - Validacion de concurrencia y produccion

## Candidato local

- Version: `50-concurrent-app-actions`.
- Punto de reversion publicado: `d6fd9f3` (`49-idempotent-daily-break`).
- Pruebas locales: 18 de 18.
- Simulacion: 100 escrituras de tareas independientes sin perdida de filas.

## Cambios que reducen contencion

- Las tareas ya no actualizan `app_data` cuando solamente cambia `task_records`.
- Comunicados, frase diaria, revision de registros y cierre de recuperaciones usan
  `POST /app/action`.
- Las configuraciones generales usan comparacion de `updated_at` y hasta seis
  intentos internos antes de responder con un error temporal reintentable.
- `PUT /state` permanece solamente para tres respaldos offline y la importacion
  administrativa. No debe usarse en acciones normales conectadas.

## Validacion posterior al despliegue

Observar Cloudflare durante al menos 15 minutos y comprobar:

1. Error 5xx menor a 1% y ningun aumento sostenido de HTTP 503.
2. Latencia p95 menor a 1.5 segundos para acciones de tarea y configuracion.
3. D1 sin crecimiento de escrituras de `app_data` durante cambios exclusivos de tareas.
4. R2 sin errores en carga directa, confirmacion o multipart.
5. Login, tarea, sustento, aprobacion, comunicado y break funcionando en PC y celular.

Si aparece perdida de estado, errores 5xx sostenidos o un flujo principal deja de
funcionar, volver temporalmente a `d6fd9f3` y conservar los registros para diagnostico.

# Task Hub - Handoff tecnico

Actualizado: 2026-08-31

## Estado actual

- Producto visible: Task Hub.
- Snapshot local activo: `publish-v23/`.
- Candidato local: `APP_VERSION = 54-access-suspended`.
- Cache PWA local: `task-hub-shell-v54-access-suspended`.
- Respuesta `/health` local: `54-access-suspended`.
- Acceso global temporalmente suspendido; D1 y R2 se conservan intactos.
- Produccion verificada antes de este cambio: `53-historical-retired-roster`, commit `ebc5269`.
- Esquema D1: `29-scoped-state-sync-1`.
- `index.html` y `operativo.html` son copias exactas.
- `functions/api/[[path]].js` y `functions/cloud/[[path]].js` son copias exactas.
- El repositorio local usa la rama `publish-final`. Verificar siempre el commit
  remoto y el despliegue de Cloudflare antes de declarar una version publicada.

## Capacidades ya presentes

- Roles de coordinador, trainer, master y observador de solo lectura.
- Desactivacion y reactivacion de personal por el coordinador de su equipo, sin borrar tareas ni sustentos.
- Retiro definitivo de personal de los listados activos, conservando su historial para reportes y auditoria.
- Segmentacion Training y Audiovisuales.
- Calendario semanal, horarios, breaks y ampliacion de jornada.
- Tareas, historial, reasignacion, revision y aprobacion.
- Sustentos obligatorios con archivos originales.
- Carga directa a R2, multipart y rutas heredadas de respaldo.
- Cola local cifrada y continuidad para tareas y evidencias sin conexion.
- Reintentos, progreso de carga y confirmacion posterior del formulario.
- Notificaciones web push y recordatorios.
- Reportes semanales Excel, ZIP y PowerPoint fotografico.
- Mantenimiento de almacenamiento y respaldo semanal estructurado.

## Almacenamiento y servicios

- `env.DB`: binding D1 obligatorio.
- `env.EVIDENCE_BUCKET`: binding R2.
- Credenciales R2 S3 para URLs firmadas: `R2_ACCOUNT_ID`, `R2_ACCESS_KEY_ID`,
  `R2_SECRET_ACCESS_KEY` y `R2_BUCKET_NAME`.
- Variables Microsoft Graph existen como integracion opcional heredada. No
  habilitar ni eliminar sin revisar datos almacenados.
- Los secretos deben permanecer en Cloudflare y nunca escribirse en el repo.

## Riesgos conocidos

1. `GET /state` arma un estado agregado y todavia puede leer muchas tareas para
   coordinadores u observadores. El historial ya se obtiene bajo demanda.
2. El frontend mantiene comprobaciones periodicas cada 60-75 segundos. La version
   condicional evita reconstruir el tablero cuando no existen cambios.
3. Frontend y backend son archivos monoliticos grandes. Los cambios deben ser
   pequenos y cubiertos por pruebas.
4. Las rutas `/api` y `/cloud` comparten implementacion y proveedor. Son fallback
   de ruta, no redundancia real frente a una caida total de Cloudflare.
5. `README.md` esta desactualizado y contiene informacion del prototipo demo. No
   usarlo para operar produccion.
6. Las versiones de frontend, cache, health y esquema no usan el mismo numero.
   No corregirlas de forma cosmetica; alinearlas en una entrega funcional futura.

## Trabajo completado en las entregas locales

- Creado `AGENTS.md` con arquitectura, reglas de seguridad, flujo de trabajo,
  pruebas y versionado.
- Creado este handoff para reducir el contexto necesario en futuras sesiones.
- Optimizado `GET /state` con version condicional. Si no existen cambios, devuelve
  `notModified` sin reconstruir ni transferir el tablero completo.
- Las lecturas de tareas e historial ahora se limitan a los propietarios visibles
  para el usuario. Un trainer consulta solo sus tareas; un coordinador, las de su
  equipo administrable; Giuliana conserva la vision global de solo lectura.
- Los recordatorios vencidos se consultan directamente desde `task_records` sin
  cargar todo el historial de tareas.
- El mantenimiento pesado ya no se programa con cada `GET /state`; continua en
  la comprobacion horaria de salud y en el estado administrativo de almacenamiento.
- La actualizacion automatica cambio de 30 segundos fijos a 60-75 segundos con
  variacion aleatoria para evitar picos simultaneos. Los guardados, reconexion y
  regreso a la aplicacion mantienen sincronizacion inmediata.
- Agregada `tests/check-state-sync-optimization.mjs`. Las 14 pruebas locales
  finalizan correctamente.
- Inicio, cambio de nombre, reprogramacion, reasignacion y recordatorios usan
  `POST /tasks/action` con `mutationId`, payload pequeno y validacion en servidor.
- Estas acciones leen solamente las tareas del propietario actual y del nuevo
  propietario cuando existe una reasignacion. El envio del tablero completo queda
  como compatibilidad para la cola offline.
- Agregada `tests/check-idempotent-task-actions.mjs`. Las 15 pruebas locales
  finalizan correctamente.
- El estado normal omite `task_history_records`. `GET /tasks/:id/history` carga
  el historial completo solamente al abrir una tarea y valida propietario, equipo,
  coordinador u observador antes de responder.
- La cola offline fusiona solo las entradas nuevas con el historial existente para
  evitar perdida o duplicacion cuando la tarea llego sin historial en el tablero.
- Agregada `tests/check-lazy-task-history.mjs`. Las 16 pruebas locales finalizan
  correctamente.
- El break por fecha usa `POST /schedule/break`, valida dia, permiso para fechas
  anteriores y cruces con tareas en el servidor, y evita una nueva escritura si
  el horario solicitado ya estaba guardado.
- El cliente bloquea el boton durante el guardado y conserva el cambio localmente
  cuando la nube falla, manteniendo `PUT /state` solo como compatibilidad offline.
- Agregada `tests/check-idempotent-daily-break.mjs`. Las 17 pruebas locales
  finalizan correctamente.
- Comunicados, frase diaria, revision de registros y cierre de recuperaciones usan
  `POST /app/action` con payload pequeno, validacion de rol e idempotencia.
- Las tareas ya no reescriben `app_data` cuando solo cambia `task_records`. Los
  cambios generales comparan `updated_at` y reintentan ante concurrencia.
- Recuperacion de clave, Info LG e interruptor de fechas anteriores modifican
  solamente `app_data` y ya no cargan todas las tareas para guardar.
- `PUT /state` queda limitado a tres respaldos offline y la importacion manual.
- Agregada `tests/check-concurrent-app-actions.mjs`, que simula 100 escrituras de
  tareas independientes. Las 18 pruebas locales finalizan correctamente.
- La guia `docs/STAGE-6-VALIDATION.md` define umbrales, observacion y rollback.
- No se modificaron usuarios, tareas ni sustentos existentes. El candidato local
  de etapa 6 aun no fue publicado.

## Comparacion de carga D1

Antes, cada consulta periodica ejecutaba lecturas de `app_data`, todas las tareas,
todo el historial y todos los usuarios, ademas de intentar mantenimiento. Ahora:

- Sin cambios: tres lecturas escalares de version en un solo `batch`, sin transferir
  tareas ni historial.
- Con cambios para trainer: `app_data` y sus tareas; el historial se carga al abrirlo.
- Con cambios para coordinador: `app_data` y tareas de su equipo; el historial se
  carga al abrirlo.
- Con cambios para Giuliana: lectura global, necesaria por su alcance de supervision.

La reduccion exacta de filas depende de los datos reales de produccion y debe
confirmarse en Cloudflare Analytics despues de publicar.

## Siguiente mejora recomendada

Publicar el candidato de etapa 6 con autorizacion explicita y aplicar la guia
`docs/STAGE-6-VALIDATION.md`. Medir errores, latencia, CPU y filas D1 con trafico
real antes de ampliar nuevas funciones. No retirar la compatibilidad offline ni
combinar esta validacion con rediseño visual, nuevos usuarios o cambios de APK.

## Criterios de aceptacion de la siguiente mejora

- Login, tareas, sustentos, aprobacion y reportes siguen funcionando.
- Un trainer no recibe datos privados de otros equipos.
- Giuliana conserva solo lectura.
- No se pierde historial ni evidencia.
- La cola offline sigue sincronizando exactamente una vez por operacion.
- Las filas leidas y el tiempo de `/state` disminuyen de forma medible.
- Todos los archivos de `tests/check-*.mjs` finalizan correctamente.

## Nota operativa

Antes de publicar, confirmar el commit realmente desplegado en Cloudflare Pages
y generar un punto de reversion. Esta carpeta local por si sola no demuestra que
coincida con produccion.

# Task Hub - Handoff tecnico

Actualizado: 2026-08-28

## Estado actual

- Producto visible: Task Hub.
- Snapshot local activo: `publish-v23/`.
- Frontend: `APP_VERSION = 46-scoped-state-sync`.
- Cache PWA: `task-hub-shell-v46-scoped-state-sync`.
- Respuesta `/health`: `46-scoped-state-sync`.
- Esquema D1: `29-scoped-state-sync-1`.
- `index.html` y `operativo.html` son copias exactas.
- `functions/api/[[path]].js` y `functions/cloud/[[path]].js` son copias exactas.
- El workspace local no contiene metadatos Git utilizables. Verificar repositorio,
  rama y commit remoto antes de cualquier despliegue.

## Capacidades ya presentes

- Roles de coordinador, trainer, master y observador de solo lectura.
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

1. `GET /state` arma un estado agregado y las funciones `loadData()` y
   `stateResponse()` pueden leer muchas tareas e historiales. Es el principal
   candidato para reducir CPU y filas leidas en D1.
2. El frontend mantiene comprobaciones periodicas cada 30 segundos. Aunque evita
   algunas ejecuciones concurrentes, todavia puede provocar lecturas repetidas.
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
- No se modificaron usuarios, tareas ni sustentos existentes. Esta entrega local
  aun no fue publicada en produccion.

## Comparacion de carga D1

Antes, cada consulta periodica ejecutaba lecturas de `app_data`, todas las tareas,
todo el historial y todos los usuarios, ademas de intentar mantenimiento. Ahora:

- Sin cambios: tres lecturas escalares de version en un solo `batch`, sin transferir
  tareas ni historial.
- Con cambios para trainer: `app_data`, sus tareas y el historial de esas tareas.
- Con cambios para coordinador: `app_data`, tareas e historial de su equipo.
- Con cambios para Giuliana: lectura global, necesaria por su alcance de supervision.

La reduccion exacta de filas depende de los datos reales de produccion y debe
confirmarse en Cloudflare Analytics despues de publicar.

## Siguiente mejora recomendada

Objetivo: reemplazar gradualmente los cambios generales mediante `PUT /state` por
operaciones pequenas y especificas para iniciar, mover, reasignar y recordar tareas.

Secuencia segura:

1. Agregar endpoints especificos e idempotentes para cada cambio de tarea.
2. Mantener temporalmente `PUT /state` como compatibilidad y cola offline.
3. Cargar el historial completo solamente al abrir una tarea concreta.
4. Medir filas leidas, escrituras, CPU y latencia en Cloudflare Analytics.
5. Probar conflictos concurrentes entre dos usuarios antes de retirar el flujo
   general de estado.

No combinar esta mejora con rediseño visual, nuevos usuarios o cambios de APK.

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

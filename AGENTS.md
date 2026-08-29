# Task Hub - Instrucciones para agentes

## Alcance

Estas instrucciones aplican a todo `publish-v23/`. Este directorio contiene la
version activa de Task Hub para web, PWA y Cloudflare Pages Functions.

Antes de modificar codigo, leer `docs/HANDOFF.md` y limitar cada trabajo a un
solo objetivo verificable.

## Arquitectura actual

- `index.html`: aplicacion principal. Contiene HTML, CSS y JavaScript.
- `operativo.html`: copia exacta de `index.html`; debe mantenerse identica.
- `sw.js`: service worker, cache PWA, continuidad y notificaciones.
- `manifest.webmanifest`: identidad e instalacion PWA.
- `functions/api/[[path]].js`: API principal de Cloudflare Pages Functions.
- `functions/cloud/[[path]].js`: copia exacta de la API; debe mantenerse identica.
- D1: usuarios, sesiones, estado, tareas, historial, notificaciones y metadatos.
- R2: archivos originales de sustentos mediante carga directa o multipart.
- `tests/`: verificaciones Node sin framework ni `package.json`.
- `vendor/`: dependencias minificadas para Excel, ZIP y PowerPoint. No editarlas
  manualmente.

`README.md` describe un prototipo antiguo. No usar sus usuarios demo, flujo ni
arquitectura como fuente de verdad para produccion.

## Reglas de seguridad

- Nunca borrar, limpiar o migrar datos D1/R2 sin autorizacion explicita, copia
  verificada y plan de reversion.
- Nunca publicar secretos, cookies, tokens, claves, variables de entorno o datos
  personales en codigo, documentacion, commits o salidas.
- No editar manualmente hashes o salts de usuarios existentes.
- No debilitar autenticacion, validacion de origen, permisos por rol ni acceso de
  solo lectura de Giuliana.
- No eliminar la cola local cifrada, reintentos, idempotencia, carga directa a R2
  ni continuidad offline para simplificar una correccion.
- Los sustentos originales deben conservar nombre, tipo, calidad, propietario,
  tarea y fecha.
- No desplegar a produccion sin autorizacion explicita del usuario.

## Reglas funcionales que deben preservarse

- Coordinadores gestionan solamente sus equipos autorizados.
- Masters administran sus propias tareas segun las reglas actuales.
- Giuliana tiene visibilidad general y no modifica informacion.
- Trainers ven y gestionan solamente el alcance permitido para su usuario.
- Estados de tarea: azul asignada, mostaza en revision, verde aprobada y rojo
  rechazada.
- La carga de sustentos debe ser reintentable y no debe perder archivos si la
  nube o la red fallan.
- Reportes Excel y PowerPoint deben conservar enlaces a originales y calidad de
  imagen.
- Horarios, breaks, extensiones de jornada y domingos deben respetar las reglas
  existentes y sus autorizaciones.

## Flujo de trabajo

1. Leer `docs/HANDOFF.md`.
2. Inspeccionar solamente los archivos relacionados con el objetivo.
3. Registrar el comportamiento actual con una prueba antes de cambiarlo.
4. Hacer el cambio minimo compatible con la arquitectura existente.
5. Si cambia `index.html`, copiar el resultado completo a `operativo.html`.
6. Si cambia `functions/api/[[path]].js`, copiar el resultado completo a
   `functions/cloud/[[path]].js`.
7. Ejecutar todas las pruebas.
8. Revisar que los pares duplicados sigan siendo identicos.
9. Actualizar `docs/HANDOFF.md` con cambios, pruebas y siguiente paso.

No mezclar cambios visuales, nuevas funciones, migraciones y optimizaciones de
backend en una misma entrega.

## Verificacion

Ejecutar todas las pruebas desde `publish-v23/`:

```powershell
Get-ChildItem -LiteralPath tests -Filter 'check-*.mjs' |
  Sort-Object Name |
  ForEach-Object { node $_.FullName }
```

Comprobar duplicados:

```powershell
Get-FileHash -Algorithm SHA256 -LiteralPath index.html, operativo.html
Get-FileHash -Algorithm SHA256 -LiteralPath `
  'functions/api/[[path]].js', 'functions/cloud/[[path]].js'
```

## Versionado

- Un cambio de frontend que deba invalidar cache requiere actualizar
  `APP_VERSION` en `index.html` y `CACHE_NAME` en `sw.js`.
- Una migracion de esquema requiere actualizar `SCHEMA_VERSION` y debe ser
  compatible con datos existentes.
- La version de `/health` debe identificar claramente la entrega desplegada.
- No cambiar versiones solo por documentacion.

## Prioridad tecnica

La lectura condicional y filtrada de D1, las acciones idempotentes y el historial
bajo demanda ya estan implementados. La siguiente prioridad es reducir los usos
restantes de `PUT /state` y medir Cloudflare Analytics. No retirar la compatibilidad
actual ni la cola offline hasta probar concurrencia, permisos, historial y sustentos.

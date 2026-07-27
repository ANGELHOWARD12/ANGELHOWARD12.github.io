const SESSION_COOKIE = "lg_session";
const SESSION_DAYS = 14;
const COORDINATOR_CODE_HASH = "e62163b1947feab8e4db70a99cffd5fb9c9f66d5e8901a4fb9775180ea780b71";

const EMPTY_DATA = {
  version: 9,
  workSettings: {
    breakStart: "12:30",
    breakEnd: "14:00"
  },
  breakSettingsByUser: {},
  registrationRequests: [],
  passwordRecoveryRequests: [],
  tasks: [],
  deletedTasks: [],
  announcements: [],
  supportRequests: [],
  dailyMotivations: []
};

const WORKDAY_START = "08:30";
const WEEKDAY_END = "19:00";
const SATURDAY_END = "14:00";
const BREAK_START = "12:30";
const BREAK_END = "14:00";
const MAX_FILE_BASE64 = 1_800_000;
const MAX_FILE_TOTAL_BASE64 = 21_000_000;
const MAX_FILE_CHUNK_BASE64 = 900_000;
const ALLOWED_FILE_EXTENSIONS = new Set([
  "jpg",
  "jpeg",
  "png",
  "webp",
  "pdf",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "ppt",
  "pptx",
  "csv",
  "txt"
]);
const ALLOWED_FILE_MIME_TYPES = new Set([
  "image/jpeg",
  "image/png",
  "image/webp",
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-powerpoint",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "application/zip",
  "application/x-zip-compressed",
  "application/msexcel",
  "application/mspowerpoint",
  "application/vnd.ms-word",
  "application/csv",
  "text/csv",
  "text/comma-separated-values",
  "text/plain",
  "application/octet-stream"
]);

export async function onRequest(context) {
  const { request, env } = context;
  if (!env.DB) return json({ ok: false, message: "La base de datos aun no esta vinculada." }, 503);

  try {
    await ensureSchema(env.DB);
    const url = new URL(request.url);
    const route = url.pathname.replace(/^\/(?:api|cloud)\/?/, "");

    if (request.method === "OPTIONS") return new Response(null, { status: 204 });
    if (!["GET", "HEAD"].includes(request.method) && !sameOrigin(request)) {
      return json({ ok: false, message: "Solicitud no permitida." }, 403);
    }

    if (route === "health" && request.method === "GET") return json({ ok: true });
    if (route === "auth/register" && request.method === "POST") return register(request, env.DB);
    if (route === "auth/login" && request.method === "POST") return login(request, env.DB);
    if (route === "auth/logout" && request.method === "POST") return logout(request, env.DB);
    if (route === "auth/recovery" && request.method === "POST") return requestRecovery(request, env.DB);

    const session = await authenticate(request, env.DB);
    if (!session) return json({ ok: false, message: "Sesion no valida." }, 401);

    if (route === "notifications/config" && request.method === "GET") return notificationConfig(env.DB);
    if (route === "notifications/subscribe" && request.method === "POST") {
      return subscribeNotifications(request, env.DB, session.user, context);
    }
    if (route === "notifications/unsubscribe" && request.method === "POST") {
      return unsubscribeNotifications(request, env.DB, session.user);
    }
    if (route === "notifications/pending" && request.method === "GET") {
      return pendingNotifications(env.DB, session.user);
    }
    if (route === "evidence/upload" && request.method === "POST") return uploadEvidence(request, env.DB, session.user);
    if (route === "evidence/upload/init" && request.method === "POST") return initEvidenceUpload(request, env.DB, session.user);
    if (route === "evidence/upload/chunk" && request.method === "POST") return uploadEvidenceChunk(request, env.DB, session.user);
    if (route === "evidence/upload/complete" && request.method === "POST") return completeEvidenceUpload(request, env.DB, session.user);
    const evidenceFileMatch = route.match(/^evidence\/([^/]+)\/(?:file|photo)$/);
    if (evidenceFileMatch && request.method === "GET") return evidenceFile(env.DB, session.user, evidenceFileMatch[1]);
    if (route === "tasks/evidence" && request.method === "POST") return submitTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/review" && request.method === "POST") return reviewTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/delete" && request.method === "POST") return deleteTaskAndArchive(request, env.DB, session.user, context);
    if (route === "state" && request.method === "GET") return getState(env.DB, session.user, context);
    if (route === "state" && request.method === "PUT") return putState(request, env.DB, session.user, context);
    if (route === "admin/users" && request.method === "POST") return createUser(request, env.DB, session.user);
    if (route === "admin/reset-password" && request.method === "POST") return resetPassword(request, env.DB, session.user);

    return json({ ok: false, message: "Ruta no encontrada." }, 404);
  } catch (error) {
    console.error("LG Task API", error);
    return json({ ok: false, message: "No se pudo completar la operacion en la nube." }, 500);
  }
}

async function ensureSchema(db) {
  await db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE COLLATE NOCASE,
      zone TEXT NOT NULL DEFAULT '',
      role TEXT NOT NULL CHECK (role IN ('Coordinador', 'Trainer')),
      status TEXT NOT NULL DEFAULT 'Activo',
      password_hash TEXT NOT NULL,
      password_salt TEXT NOT NULL,
      created_at INTEGER NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS sessions (
      token_hash TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      expires_at INTEGER NOT NULL,
      created_at INTEGER NOT NULL,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS app_data (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      data TEXT NOT NULL,
      updated_at INTEGER NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS evidence_files (
      id TEXT PRIMARY KEY,
      task_id TEXT NOT NULL,
      owner_id TEXT NOT NULL,
      submitted_by_id TEXT NOT NULL,
      file_name TEXT NOT NULL,
      mime_type TEXT NOT NULL,
      photo_base64 TEXT NOT NULL,
      created_at INTEGER NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS evidence_file_chunks (
      file_id TEXT NOT NULL,
      chunk_index INTEGER NOT NULL,
      chunk_base64 TEXT NOT NULL,
      PRIMARY KEY (file_id, chunk_index),
      FOREIGN KEY (file_id) REFERENCES evidence_files(id) ON DELETE CASCADE
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS notification_subscriptions (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      endpoint TEXT NOT NULL UNIQUE,
      p256dh TEXT NOT NULL DEFAULT '',
      auth TEXT NOT NULL DEFAULT '',
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS user_notifications (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      title TEXT NOT NULL,
      body TEXT NOT NULL,
      target_url TEXT NOT NULL DEFAULT '/',
      source_key TEXT NOT NULL UNIQUE,
      created_at INTEGER NOT NULL,
      delivered_at INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at INTEGER NOT NULL
    )`),
    db.prepare("INSERT OR IGNORE INTO app_data (id, data, updated_at) VALUES (1, ?, 0)").bind(JSON.stringify(EMPTY_DATA)),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sessions_expiry ON sessions(expires_at)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_task ON evidence_files(task_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_submitter ON evidence_files(submitted_by_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_chunks_file ON evidence_file_chunks(file_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_notification_subscriptions_user ON notification_subscriptions(user_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_user_notifications_pending ON user_notifications(user_id, delivered_at)")
  ]);
}

async function register(request, db) {
  const body = await readJson(request);
  const name = clean(body.name);
  const email = clean(body.email).toLowerCase();
  const zone = clean(body.zone);
  const role = body.role === "Coordinador" ? "Coordinador" : "Trainer";
  const password = String(body.password || "");

  if (name.length < 3 || !validEmail(email) || password.length < 6) {
    return json({ ok: false, message: "Completa nombre, correo valido y una clave de minimo 6 caracteres." }, 400);
  }
  if (role === "Coordinador" && (await sha256Hex(clean(body.coordinatorCode))) !== COORDINATOR_CODE_HASH) {
    return json({ ok: false, message: "Codigo de coordinador incorrecto." }, 403);
  }

  const existing = await db.prepare("SELECT id FROM users WHERE email = ?").bind(email).first();
  if (existing) return json({ ok: false, message: "Ese correo ya esta registrado." }, 409);

  const passwordData = await hashPassword(password);
  const user = {
    id: crypto.randomUUID(),
    name,
    email,
    zone,
    role,
    status: "Activo",
    createdAt: Date.now()
  };
  await db
    .prepare("INSERT INTO users (id, name, email, zone, role, status, password_hash, password_salt, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
    .bind(user.id, name, email, zone, role, user.status, passwordData.hash, passwordData.salt, user.createdAt)
    .run();

  return json({ ok: true, user: publicUser(user) }, 201);
}

async function login(request, db) {
  const body = await readJson(request);
  const email = clean(body.email).toLowerCase();
  const password = String(body.password || "");
  const row = await db.prepare("SELECT * FROM users WHERE email = ?").bind(email).first();
  if (!row || row.status !== "Activo" || !(await verifyPassword(password, row.password_salt, row.password_hash))) {
    return json({ ok: false, message: "Correo o clave incorrectos, o usuario no aprobado." }, 401);
  }

  const token = randomToken(32);
  const tokenHash = await sha256Hex(token);
  const expiresAt = Date.now() + SESSION_DAYS * 86400000;
  await db.prepare("DELETE FROM sessions WHERE expires_at < ?").bind(Date.now()).run();
  await db
    .prepare("INSERT INTO sessions (token_hash, user_id, expires_at, created_at) VALUES (?, ?, ?, ?)")
    .bind(tokenHash, row.id, expiresAt, Date.now())
    .run();

  const response = await stateResponse(db, row);
  response.headers.append("Set-Cookie", cookieValue(token, SESSION_DAYS * 86400));
  return response;
}

async function logout(request, db) {
  const token = readCookie(request, SESSION_COOKIE);
  if (token) await db.prepare("DELETE FROM sessions WHERE token_hash = ?").bind(await sha256Hex(token)).run();
  const response = json({ ok: true });
  response.headers.append("Set-Cookie", cookieValue("", 0));
  return response;
}

async function requestRecovery(request, db) {
  const body = await readJson(request);
  const email = clean(body.email).toLowerCase();
  const user = validEmail(email) ? await db.prepare("SELECT id, email FROM users WHERE email = ?").bind(email).first() : null;
  if (user) {
    const data = await loadData(db);
    const pending = data.passwordRecoveryRequests.some((item) => item.userId === user.id && item.status === "Pendiente");
    if (!pending) {
      data.passwordRecoveryRequests.unshift({
        id: crypto.randomUUID(),
        userId: user.id,
        email: user.email,
        status: "Pendiente",
        createdAt: Date.now()
      });
      await saveData(db, data);
    }
  }
  return json({ ok: true, message: "Si el correo esta registrado, el coordinador recibira la solicitud." });
}

async function authenticate(request, db) {
  const token = readCookie(request, SESSION_COOKIE);
  if (!token) return null;
  const row = await db
    .prepare(`SELECT u.*, s.expires_at FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token_hash = ?`)
    .bind(await sha256Hex(token))
    .first();
  if (!row || row.expires_at < Date.now() || row.status !== "Activo") return null;
  return { user: row };
}

async function getState(db, user, context) {
  const data = await loadData(db);
  context.waitUntil(dispatchDueReminders(db, data));
  return stateResponse(db, user, data);
}

async function notificationConfig(db) {
  const keys = await ensureVapidKeys(db);
  return json({ ok: true, publicKey: keys.publicKey });
}

async function subscribeNotifications(request, db, user, context) {
  const body = await readJson(request);
  const endpoint = clean(body.endpoint);
  const p256dh = clean(body.keys?.p256dh);
  const auth = clean(body.keys?.auth);
  if (!endpoint.startsWith("https://") || endpoint.length > 2000) {
    return json({ ok: false, message: "La suscripcion de notificaciones no es valida." }, 400);
  }
  const subscriptionId = await sha256Hex(endpoint);
  const now = Date.now();
  await db
    .prepare(
      `INSERT INTO notification_subscriptions (id, user_id, endpoint, p256dh, auth, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(endpoint) DO UPDATE SET user_id = excluded.user_id, p256dh = excluded.p256dh,
       auth = excluded.auth, updated_at = excluded.updated_at`
    )
    .bind(subscriptionId, user.id, endpoint, p256dh, auth, now, now)
    .run();
  context.waitUntil(pushNotificationsForUsers(db, [user.id]));
  return json({ ok: true });
}

async function unsubscribeNotifications(request, db, user) {
  const body = await readJson(request);
  const endpoint = clean(body.endpoint);
  if (endpoint) {
    await db.prepare("DELETE FROM notification_subscriptions WHERE endpoint = ? AND user_id = ?").bind(endpoint, user.id).run();
  }
  return json({ ok: true });
}

async function pendingNotifications(db, user) {
  const rows = await db
    .prepare(
      `SELECT id, title, body, target_url, created_at
       FROM user_notifications WHERE user_id = ? AND delivered_at = 0
       ORDER BY created_at ASC LIMIT 20`
    )
    .bind(user.id)
    .all();
  const notifications = (rows.results || []).map((row) => ({
    id: row.id,
    title: row.title,
    body: row.body,
    url: row.target_url,
    createdAt: row.created_at
  }));
  if (notifications.length) {
    await db.batch(
      notifications.map((item) =>
        db.prepare("UPDATE user_notifications SET delivered_at = ? WHERE id = ? AND user_id = ?").bind(Date.now(), item.id, user.id)
      )
    );
  }
  return json({ ok: true, notifications });
}

async function uploadEvidence(request, db, user) {
  const body = await readJson(request, 2_000_000);
  const taskId = clean(body.taskId);
  const data = await loadData(db);
  const task = data.tasks.find((item) => item.id === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (user.role !== "Coordinador" && task.ownerId !== user.id) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }

  const fileId = crypto.randomUUID();
  const fileName = clean(body.fileName).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "archivo-sustento";
  const extension = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
  const fileMatch = /^data:([^;,]{1,150});base64,([A-Za-z0-9+/=]+)$/.exec(String(body.fileData || body.photoData || ""));
  const mimeType = clean(fileMatch?.[1]).toLowerCase();
  if (!fileMatch || !ALLOWED_FILE_EXTENSIONS.has(extension) || !ALLOWED_FILE_MIME_TYPES.has(mimeType)) {
    return json({ ok: false, message: "El archivo no tiene un formato permitido." }, 400);
  }
  if (fileMatch[2].length > MAX_FILE_BASE64) {
    return json({ ok: false, message: "El archivo supera el limite de 1.35 MB permitido por la nube." }, 413);
  }

  const createdAt = Date.now();
  await db
    .prepare(
      "INSERT INTO evidence_files (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    )
    .bind(fileId, task.id, task.ownerId, user.id, fileName, mimeType, fileMatch[2], createdAt)
    .run();

  return json({
    ok: true,
    file: { id: fileId, name: fileName, mimeType, createdAt, url: `/cloud/evidence/${fileId}/file` }
  }, 201);
}

async function initEvidenceUpload(request, db, user) {
  const body = await readJson(request);
  const taskId = clean(body.taskId);
  const data = await loadData(db);
  const task = data.tasks.find((item) => item.id === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (user.role !== "Coordinador" && task.ownerId !== user.id) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }

  const fileName = clean(body.fileName).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "archivo-sustento";
  const extension = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
  const mimeType = clean(body.mimeType).toLowerCase() || mimeTypeForExtension(extension);
  const totalBase64 = Number(body.totalBase64 || 0);
  const chunkCount = Number(body.chunkCount || 0);
  if (
    !ALLOWED_FILE_EXTENSIONS.has(extension) ||
    !ALLOWED_FILE_MIME_TYPES.has(mimeType) ||
    !Number.isInteger(chunkCount) ||
    chunkCount < 1 ||
    chunkCount > 30 ||
    totalBase64 < 1 ||
    totalBase64 > MAX_FILE_TOTAL_BASE64
  ) {
    return json({ ok: false, message: "El archivo no tiene un formato o tamano permitido." }, 400);
  }

  const fileId = crypto.randomUUID();
  const createdAt = Date.now();
  await db
    .prepare(
      "INSERT INTO evidence_files (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at) VALUES (?, ?, ?, ?, ?, ?, '', ?)"
    )
    .bind(fileId, task.id, task.ownerId, user.id, fileName, mimeType, createdAt)
    .run();
  return json({ ok: true, fileId, chunkCount, createdAt }, 201);
}

async function uploadEvidenceChunk(request, db, user) {
  const body = await readJson(request, 1_050_000);
  const fileId = clean(body.fileId);
  const chunkIndex = Number(body.chunkIndex);
  const chunkBase64 = String(body.chunkBase64 || "");
  const row = await db.prepare("SELECT owner_id, submitted_by_id FROM evidence_files WHERE id = ?").bind(fileId).first();
  if (!row) return json({ ok: false, message: "La carga ya no existe." }, 404);
  if (user.role !== "Coordinador" && row.submitted_by_id !== user.id && row.owner_id !== user.id) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  if (
    !Number.isInteger(chunkIndex) ||
    chunkIndex < 0 ||
    chunkIndex > 29 ||
    !chunkBase64 ||
    chunkBase64.length > MAX_FILE_CHUNK_BASE64 ||
    !/^[A-Za-z0-9+/=]+$/.test(chunkBase64)
  ) {
    return json({ ok: false, message: "Uno de los bloques del archivo no es valido." }, 400);
  }
  await db
    .prepare("INSERT OR REPLACE INTO evidence_file_chunks (file_id, chunk_index, chunk_base64) VALUES (?, ?, ?)")
    .bind(fileId, chunkIndex, chunkBase64)
    .run();
  return json({ ok: true, chunkIndex });
}

async function completeEvidenceUpload(request, db, user) {
  const body = await readJson(request);
  const fileId = clean(body.fileId);
  const expectedChunks = Number(body.chunkCount);
  const row = await db.prepare("SELECT * FROM evidence_files WHERE id = ?").bind(fileId).first();
  if (!row) return json({ ok: false, message: "La carga ya no existe." }, 404);
  if (user.role !== "Coordinador" && row.submitted_by_id !== user.id && row.owner_id !== user.id) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const summary = await db
    .prepare("SELECT COUNT(*) AS chunks, COALESCE(SUM(LENGTH(chunk_base64)), 0) AS total FROM evidence_file_chunks WHERE file_id = ?")
    .bind(fileId)
    .first();
  if (
    !Number.isInteger(expectedChunks) ||
    expectedChunks < 1 ||
    Number(summary?.chunks || 0) !== expectedChunks ||
    Number(summary?.total || 0) > MAX_FILE_TOTAL_BASE64
  ) {
    return json({ ok: false, message: "Faltan bloques del archivo. Intenta subirlo nuevamente." }, 400);
  }
  return json({
    ok: true,
    file: {
      id: row.id,
      name: row.file_name,
      mimeType: row.mime_type,
      createdAt: row.created_at,
      url: `/cloud/evidence/${row.id}/file`
    }
  }, 201);
}

async function evidenceFile(db, user, fileId) {
  const row = await db.prepare("SELECT * FROM evidence_files WHERE id = ?").bind(clean(fileId)).first();
  if (!row) return json({ ok: false, message: "Archivo no encontrado." }, 404);
  const allowed = user.role === "Coordinador" || row.submitted_by_id === user.id || row.owner_id === user.id;
  if (!allowed) return json({ ok: false, message: "No tienes acceso a este archivo." }, 403);

  let fileBase64 = row.photo_base64 || "";
  if (!fileBase64) {
    const chunks = await db
      .prepare("SELECT chunk_base64 FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC")
      .bind(row.id)
      .all();
    fileBase64 = (chunks.results || []).map((item) => item.chunk_base64).join("");
  }
  if (!fileBase64) return json({ ok: false, message: "El archivo esta incompleto." }, 409);
  const binary = atob(fileBase64);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const safeName = String(row.file_name || "sustento.jpg").replace(/["\r\n]/g, "-");
  const disposition = String(row.mime_type).startsWith("image/") || row.mime_type === "application/pdf" ? "inline" : "attachment";
  return new Response(bytes, {
    headers: {
      "Content-Type": row.mime_type,
      "Content-Disposition": `${disposition}; filename="${safeName}"`,
      "Cache-Control": "private, max-age=300",
      "Content-Length": String(bytes.length),
      "X-Content-Type-Options": "nosniff"
    }
  });
}

async function submitTaskEvidence(request, db, user, context) {
  const body = await readJson(request, 1_000_000);
  const taskId = clean(body.taskId);
  const submittedEvidence = body.evidence || {};
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (clean(task.ownerId) !== user.id) {
    return json({ ok: false, message: "Solo el responsable de la tarea puede enviar el sustento." }, 403);
  }
  if (task.status === "Cumplida") {
    return json({ ok: false, message: "La tarea ya fue aprobada y completada." }, 409);
  }

  const result = clean(submittedEvidence.result);
  const notes = clean(submittedEvidence.notes).slice(0, 3000);
  const store = clean(submittedEvidence.store).slice(0, 180);
  const product = clean(submittedEvidence.product).slice(0, 180);
  if (result !== "Realizado" || !store) {
    return json({ ok: false, message: "Completa la tienda. El resultado debe ser Realizado." }, 400);
  }

  const submittedFiles = Array.isArray(submittedEvidence.files) ? submittedEvidence.files.slice(0, 8) : [];
  if (!submittedFiles.length) {
    return json({ ok: false, message: "Debes adjuntar al menos una foto o archivo de sustento." }, 400);
  }
  const files = [];
  for (const submittedFile of submittedFiles) {
    const fileId = clean(submittedFile.id);
    const file = await db
      .prepare(
        "SELECT id, task_id, submitted_by_id, file_name, mime_type FROM evidence_files WHERE id = ?"
      )
      .bind(fileId)
      .first();
    if (!file || clean(file.task_id) !== taskId || clean(file.submitted_by_id) !== user.id) {
      return json({ ok: false, message: "Uno de los archivos no pertenece a esta tarea." }, 403);
    }
    files.push({
      id: file.id,
      name: file.file_name,
      mimeType: file.mime_type
    });
  }

  const evidenceId = clean(submittedEvidence.id) || crypto.randomUUID();
  task.evidence = Array.isArray(task.evidence) ? task.evidence : [];
  if (task.evidence.some((item) => clean(item.id) === evidenceId)) {
    return stateResponse(db, user, data);
  }
  const submittedAt = Date.now();
  const evidence = {
    id: evidenceId,
    submittedById: user.id,
    submittedAt,
    store,
    product,
    result,
    notes,
    files,
    cloudPath: clean(submittedEvidence.cloudPath).slice(0, 800),
    review: "Pendiente"
  };
  task.evidence.push(evidence);
  task.history = Array.isArray(task.history) ? task.history : [];
  task.status = "En revision";
  task.blockedReason = "";
  task.blockedAt = 0;
  task.history.push({
    type: "Sustento enviado",
    fromId: user.id,
    reason: `${files.length} archivo${files.length === 1 ? "" : "s"} enviado${files.length === 1 ? "" : "s"}`,
    at: submittedAt
  });

  await saveData(db, data);
  const coordinators = await db
    .prepare("SELECT id FROM users WHERE role = 'Coordinador' AND status = 'Activo'")
    .all();
  const notifications = (coordinators.results || [])
    .filter((coordinator) => clean(coordinator.id) !== user.id)
    .map((coordinator) => ({
      userId: coordinator.id,
      title: "Nuevo sustento por revisar",
      body: `${clean(task.title)} | ${clean(user.name)}`,
      url: `/?view=evidenceView&task=${encodeURIComponent(taskId)}`,
      sourceKey: `evidence:${taskId}:${evidenceId}:${clean(coordinator.id)}`
    }));
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function reviewTaskEvidence(request, db, user, context) {
  if (user.role !== "Coordinador") {
    return json({ ok: false, message: "Solo el coordinador puede aprobar o rechazar sustentos." }, 403);
  }
  const body = await readJson(request);
  const taskId = clean(body.taskId);
  const status = clean(body.status);
  if (!["Cumplida", "Observada"].includes(status)) {
    return json({ ok: false, message: "La revision indicada no es valida." }, 400);
  }
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  const latestEvidence = Array.isArray(task.evidence) ? task.evidence.at(-1) : null;
  if (!latestEvidence) {
    return json({ ok: false, message: "La tarea aun no tiene un sustento para revisar." }, 409);
  }

  const reviewedAt = Date.now();
  task.status = status;
  latestEvidence.review = status === "Cumplida" ? "Aprobado" : "Rechazado";
  latestEvidence.reviewedById = user.id;
  latestEvidence.reviewedAt = reviewedAt;
  task.history = Array.isArray(task.history) ? task.history : [];
  task.history.push({
    type: status === "Cumplida" ? "Aprobacion" : "Rechazo",
    byId: user.id,
    reason: status === "Cumplida" ? "Sustento aprobado y tarea completada" : "Sustento rechazado por el coordinador",
    at: reviewedAt
  });
  await saveData(db, data);

  const notification = {
    userId: task.ownerId,
    title: status === "Cumplida" ? "Tarea aprobada y completada" : "Sustento rechazado",
    body: clean(task.title),
    url: `/?view=tasksView&task=${encodeURIComponent(taskId)}`,
    sourceKey: `review:${taskId}:${status}:${reviewedAt}`
  };
  const notifiedUsers = await queueNotifications(db, [notification]);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

function archiveDeletedTask(data, task, actor, reason = "") {
  data.deletedTasks = Array.isArray(data.deletedTasks) ? data.deletedTasks : [];
  const deletedAt = Date.now();
  const record = {
    id: crypto.randomUUID(),
    originalTaskId: clean(task.id),
    title: clean(task.title),
    ownerId: clean(task.ownerId),
    createdById: clean(task.createdById),
    category: normalizeTaskCategory(task.category),
    priority: clean(task.priority),
    dueDate: clean(task.dueDate),
    startTime: clean(task.startTime) || WORKDAY_START,
    endTime: clean(task.endTime) || "09:30",
    product: clean(task.product),
    description: clean(task.description),
    status: clean(task.status),
    evidenceCount: Array.isArray(task.evidence) ? task.evidence.length : 0,
    deletedById: clean(actor.id),
    deletedByName: clean(actor.name) || "Coordinador",
    deletedAt,
    reason: clean(reason).slice(0, 1000) || "Tarea retirada de la agenda por el coordinador."
  };
  data.deletedTasks.unshift(record);
  data.deletedTasks = data.deletedTasks.slice(0, 500);
  return record;
}

async function deleteTaskEvidenceFiles(db, taskId) {
  await db
    .prepare("DELETE FROM evidence_file_chunks WHERE file_id IN (SELECT id FROM evidence_files WHERE task_id = ?)")
    .bind(taskId)
    .run();
  await db.prepare("DELETE FROM evidence_files WHERE task_id = ?").bind(taskId).run();
}

async function deleteTaskAndArchive(request, db, user, context) {
  if (user.role !== "Coordinador") {
    return json({ ok: false, message: "Solo el coordinador puede eliminar tareas." }, 403);
  }
  const body = await readJson(request);
  const taskId = clean(body.taskId);
  const reason = clean(body.reason);
  if (!taskId || !reason) {
    return json({ ok: false, message: "Indica la tarea y el motivo de la eliminacion." }, 400);
  }
  const data = await loadData(db);
  const taskIndex = data.tasks.findIndex((task) => clean(task.id) === taskId);
  if (taskIndex < 0) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  const task = data.tasks[taskIndex];
  const archived = archiveDeletedTask(data, task, user, reason);
  data.tasks.splice(taskIndex, 1);
  await deleteTaskEvidenceFiles(db, taskId);
  await saveData(db, data);

  const notifications = task.ownerId === user.id
    ? []
    : [{
        userId: task.ownerId,
        title: "Tarea eliminada por el coordinador",
        body: `${clean(task.title)} | ${clean(task.dueDate)} ${clean(task.startTime)}-${clean(task.endTime)}`,
        url: "/?view=tasksView",
        sourceKey: `task-deleted:${taskId}:${archived.deletedAt}`
      }];
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function stateResponse(db, user, loadedData = null) {
  const data = loadedData || (await loadData(db));
  const coordinator = user.role === "Coordinador";
  const userRows = coordinator
    ? await db.prepare("SELECT id, name, email, zone, role, status, created_at FROM users ORDER BY created_at ASC").all()
    : await db
        .prepare("SELECT id, name, email, zone, role, status, created_at FROM users WHERE status = 'Activo' ORDER BY role ASC, name ASC")
        .all();
  const users = (userRows.results || []).map(publicUser);
  const state = {
    ...EMPTY_DATA,
    ...data,
    activeUserId: user.id,
    users,
    tasks: coordinator ? data.tasks : data.tasks.filter((task) => task.ownerId === user.id),
    deletedTasks: coordinator
      ? data.deletedTasks
      : (data.deletedTasks || []).filter((task) => task.ownerId === user.id),
    announcements: coordinator
      ? data.announcements
      : data.announcements.filter((item) => item.audience === "all" || item.targetId === user.id),
    supportRequests: coordinator
      ? data.supportRequests
      : data.supportRequests.filter((item) => item.fromId === user.id || item.toId === user.id),
    passwordRecoveryRequests: coordinator ? data.passwordRecoveryRequests : [],
    registrationRequests: coordinator ? data.registrationRequests : []
  };
  return json({ ok: true, state, user: publicUser(user) });
}

async function putState(request, db, user, context) {
  const body = await readJson(request, 6_000_000);
  const submitted = body.state || {};
  const current = await loadData(db);
  const notifications = [];
  const personalBreak = submitted.breakSettingsByUser?.[user.id];
  if (validWorkSettings(personalBreak)) {
    current.breakSettingsByUser = {
      ...(current.breakSettingsByUser || {}),
      [user.id]: normalizeWorkSettings(personalBreak)
    };
  }

  if (user.role === "Coordinador") {
    for (const key of ["registrationRequests", "passwordRecoveryRequests", "announcements", "supportRequests", "dailyMotivations"]) {
      if (Array.isArray(submitted[key])) current[key] = submitted[key];
    }
    if (Array.isArray(submitted.tasks)) {
      const previousTasks = new Map(current.tasks.map((task) => [clean(task.id), task]));
      const submittedTaskIds = new Set(submitted.tasks.map((task) => clean(task.id)).filter(Boolean));
      const removedTasks = current.tasks.filter((task) => clean(task.id) && !submittedTaskIds.has(clean(task.id)));
      current.tasks = submitted.tasks
        .map((task) => {
          const normalizedTask = {
            ...task,
            category: normalizeTaskCategory(task.category)
          };
          const previous = previousTasks.get(clean(normalizedTask.id));
          const ownerOrScheduleChanged =
            !previous ||
            clean(previous.ownerId) !== clean(normalizedTask.ownerId) ||
            previous.dueDate !== normalizedTask.dueDate ||
            previous.startTime !== normalizedTask.startTime ||
            previous.endTime !== normalizedTask.endTime;
          if (
            ownerOrScheduleChanged &&
            !validWorkSchedule(
              clean(normalizedTask.dueDate),
              clean(normalizedTask.startTime),
              clean(normalizedTask.endTime),
              breakSettingsForUser(current, normalizedTask.ownerId)
            )
          ) {
            return previous || null;
          }
          return normalizedTask;
        })
        .filter(Boolean);
      for (const task of current.tasks) {
        const previous = previousTasks.get(clean(task.id));
        const assignedNow = !previous && clean(task.ownerId);
        const reassignedNow = previous && clean(previous.ownerId) !== clean(task.ownerId);
        const scheduleChanged =
          previous &&
          (previous.dueDate !== task.dueDate || previous.startTime !== task.startTime || previous.endTime !== task.endTime);
        if ((assignedNow || reassignedNow || scheduleChanged) && task.ownerId !== user.id) {
          const action = assignedNow ? "Nueva tarea asignada" : reassignedNow ? "Tarea reasignada" : "Horario actualizado";
          const marker = task.history?.at(-1)?.at || task.createdAt || Date.now();
          notifications.push({
            userId: task.ownerId,
            title: action,
            body: `${clean(task.title)} | ${clean(task.dueDate)} ${clean(task.startTime)}-${clean(task.endTime)}`,
            url: `/?view=tasksView&task=${encodeURIComponent(clean(task.id))}`,
            sourceKey: `task:${clean(task.id)}:${clean(task.ownerId)}:${action}:${marker}`
          });
        }
      }
      for (const removedTask of removedTasks) {
        archiveDeletedTask(current, removedTask, user, "Tarea retirada de la agenda por el coordinador.");
        await deleteTaskEvidenceFiles(db, clean(removedTask.id));
      }
    }
  } else {
    const submittedTasks = new Map((submitted.tasks || []).map((task) => [task.id, task]));
    const activeTrainerRows = await db.prepare("SELECT id FROM users WHERE role = 'Trainer' AND status = 'Activo'").all();
    const activeTrainerIds = new Set((activeTrainerRows.results || []).map((row) => row.id));
    current.tasks = current.tasks.map((task) => {
      if (task.ownerId !== user.id) return task;
      const next = submittedTasks.get(task.id);
      if (!next) return task;
      const requestedOwnerId = clean(next.ownerId);
      const ownerChanged = requestedOwnerId && requestedOwnerId !== task.ownerId && activeTrainerIds.has(requestedOwnerId);
      const nextHistory = Array.isArray(next.history) ? next.history : task.history;
      const lastHistory = nextHistory?.at(-1);
      const validReassignment =
        ownerChanged &&
        lastHistory?.type === "Reasignacion" &&
        lastHistory.fromId === user.id &&
        lastHistory.toId === requestedOwnerId &&
        lastHistory.byId === user.id &&
        clean(lastHistory.reason) &&
        validWorkSchedule(
          clean(task.dueDate),
          clean(task.startTime),
          clean(task.endTime),
          breakSettingsForUser(current, requestedOwnerId)
        ) &&
        !hasTaskConflict(
          current.tasks,
          requestedOwnerId,
          clean(task.dueDate),
          clean(task.startTime),
          clean(task.endTime),
          clean(task.id)
        );
      const requestedDate = clean(next.dueDate);
      const requestedStart = clean(next.startTime);
      const requestedEnd = clean(next.endTime);
      const scheduleChanged =
        requestedDate !== task.dueDate ||
        requestedStart !== (task.startTime || WORKDAY_START) ||
        requestedEnd !== (task.endTime || "09:30");
      const validReschedule =
        !ownerChanged &&
        scheduleChanged &&
        task.status !== "Cumplida" &&
        lastHistory?.type === "Reprogramacion" &&
        lastHistory.byId === user.id &&
        lastHistory.fromDate === task.dueDate &&
        lastHistory.toDate === requestedDate &&
        lastHistory.fromStartTime === (task.startTime || WORKDAY_START) &&
        lastHistory.fromEndTime === (task.endTime || "09:30") &&
        lastHistory.toStartTime === requestedStart &&
        lastHistory.toEndTime === requestedEnd &&
        clean(lastHistory.reason) &&
        validWorkSchedule(requestedDate, requestedStart, requestedEnd, breakSettingsForUser(current, user.id)) &&
        !hasTaskConflict(current.tasks, user.id, requestedDate, requestedStart, requestedEnd, task.id);
      if (validReassignment) {
        notifications.push({
          userId: requestedOwnerId,
          title: "Tarea reasignada",
          body: `${clean(task.title)} | ${clean(task.dueDate)} ${clean(task.startTime)}-${clean(task.endTime)}`,
          url: `/?view=tasksView&task=${encodeURIComponent(clean(task.id))}`,
          sourceKey: `task:${clean(task.id)}:${requestedOwnerId}:reassigned:${lastHistory.at || Date.now()}`
        });
      }
      const requestedStatus = clean(next.status);
      const validStart = task.status === "Pendiente" && requestedStatus === "En proceso";
      return {
        ...task,
        ownerId: validReassignment ? requestedOwnerId : task.ownerId,
        dueDate: validReschedule ? requestedDate : task.dueDate,
        startTime: validReschedule ? requestedStart : task.startTime,
        endTime: validReschedule ? requestedEnd : task.endTime,
        status: validReassignment ? "Pendiente" : validStart ? "En proceso" : task.status,
        evidence: task.evidence,
        reminders: validReassignment
          ? []
          : validReminderList(next.reminders, user.id, task.ownerId)
            ? next.reminders
            : task.reminders || [],
        history:
          validReassignment || validReschedule
            ? nextHistory
            : task.history,
        blockedReason: validReassignment ? "" : task.blockedReason,
        blockedAt: validReassignment ? 0 : task.blockedAt
      };
    });

    const knownTaskIds = new Set(current.tasks.map((task) => task.id));
    for (const task of submitted.tasks || []) {
      const taskId = clean(task.id);
      if (!taskId || knownTaskIds.has(taskId)) continue;
      const dueDate = clean(task.dueDate);
      const startTime = clean(task.startTime);
      const endTime = clean(task.endTime);
      if (
        task.ownerId !== user.id ||
        task.createdById !== user.id ||
        clean(task.title).length < 2 ||
        !validWorkSchedule(dueDate, startTime, endTime, breakSettingsForUser(current, user.id)) ||
        hasTaskConflict(current.tasks, user.id, dueDate, startTime, endTime)
      ) {
        continue;
      }
      const createdAt = Date.now();
      current.tasks.unshift({
        id: taskId,
        title: clean(task.title).slice(0, 140),
        ownerId: user.id,
        createdById: user.id,
        category: normalizeTaskCategory(task.category),
        priority: ["Alta", "Media", "Baja"].includes(task.priority) ? task.priority : "Media",
        dueDate,
        startTime,
        endTime,
        product: clean(task.product).slice(0, 120),
        description: clean(task.description).slice(0, 1500),
        status: "Pendiente",
        createdAt,
        history: [
          {
            type: "Asignacion",
            toId: user.id,
            byId: user.id,
            reason: "Tarea personal creada por el trainer",
            at: createdAt
          }
        ],
        evidence: [],
        reminders: []
      });
      knownTaskIds.add(taskId);
    }
  }

  await saveData(db, current);
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  context.waitUntil(dispatchDueReminders(db, current));
  return stateResponse(db, user, current);
}

function validReminderList(reminders, actorId, ownerId) {
  if (!Array.isArray(reminders) || actorId !== ownerId || reminders.length > 30) return false;
  return reminders.every(
    (reminder) =>
      clean(reminder.id) &&
      clean(reminder.userId) === actorId &&
      Number.isFinite(Number(reminder.at)) &&
      Number(reminder.at) > 0
  );
}

async function queueNotifications(db, notifications) {
  const notifiedUsers = new Set();
  for (const item of notifications) {
    if (!clean(item.userId) || !clean(item.sourceKey)) continue;
    const result = await db
      .prepare(
        `INSERT OR IGNORE INTO user_notifications
         (id, user_id, title, body, target_url, source_key, created_at, delivered_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, 0)`
      )
      .bind(
        crypto.randomUUID(),
        clean(item.userId),
        clean(item.title).slice(0, 100) || "LGTASK",
        clean(item.body).slice(0, 280),
        clean(item.url).slice(0, 500) || "/",
        clean(item.sourceKey).slice(0, 500),
        Date.now()
      )
      .run();
    if (Number(result?.meta?.changes || 0) > 0) notifiedUsers.add(clean(item.userId));
  }
  return Array.from(notifiedUsers);
}

async function dispatchDueReminders(db, data) {
  const now = Date.now();
  const notifications = [];
  for (const task of data.tasks || []) {
    if (task.status === "Cumplida") continue;
    for (const reminder of task.reminders || []) {
      const reminderAt = Number(reminder.at || 0);
      const userId = clean(reminder.userId || task.ownerId);
      if (!userId || !reminderAt || reminderAt > now) continue;
      notifications.push({
        userId,
        title: "Recordatorio de tarea",
        body: `${clean(task.title)} | ${clean(task.dueDate)} ${clean(task.startTime)}-${clean(task.endTime)}`,
        url: `/?view=tasksView&task=${encodeURIComponent(clean(task.id))}`,
        sourceKey: `reminder:${clean(task.id)}:${clean(reminder.id)}`
      });
    }
  }
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) await pushNotificationsForUsers(db, notifiedUsers);
}

async function ensureVapidKeys(db) {
  const row = await db.prepare("SELECT value FROM app_settings WHERE key = 'vapid_keys'").first();
  if (row?.value) {
    try {
      return JSON.parse(row.value);
    } catch {
      // Regenerate invalid configuration.
    }
  }
  const pair = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, true, ["sign", "verify"]);
  const privateJwk = await crypto.subtle.exportKey("jwk", pair.privateKey);
  const publicJwk = await crypto.subtle.exportKey("jwk", pair.publicKey);
  const x = fromBase64Url(publicJwk.x);
  const y = fromBase64Url(publicJwk.y);
  const publicBytes = new Uint8Array(65);
  publicBytes[0] = 4;
  publicBytes.set(x, 1);
  publicBytes.set(y, 33);
  const keys = {
    publicKey: toBase64Url(publicBytes),
    privateJwk
  };
  await db
    .prepare(
      `INSERT INTO app_settings (key, value, updated_at) VALUES ('vapid_keys', ?, ?)
       ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at`
    )
    .bind(JSON.stringify(keys), Date.now())
    .run();
  return keys;
}

async function pushNotificationsForUsers(db, userIds) {
  const uniqueUserIds = Array.from(new Set(userIds.map(clean).filter(Boolean)));
  if (!uniqueUserIds.length) return;
  const keys = await ensureVapidKeys(db);
  for (const userId of uniqueUserIds) {
    const rows = await db.prepare("SELECT id, endpoint FROM notification_subscriptions WHERE user_id = ?").bind(userId).all();
    for (const subscription of rows.results || []) {
      try {
        const response = await sendWebPush(subscription.endpoint, keys);
        if (response.status === 404 || response.status === 410) {
          await db.prepare("DELETE FROM notification_subscriptions WHERE id = ?").bind(subscription.id).run();
        }
      } catch (error) {
        console.error("LGTASK push", error);
      }
    }
  }
}

async function sendWebPush(endpoint, keys) {
  const audience = new URL(endpoint).origin;
  const header = toBase64Url(new TextEncoder().encode(JSON.stringify({ typ: "JWT", alg: "ES256" })));
  const payload = toBase64Url(
    new TextEncoder().encode(
      JSON.stringify({
        aud: audience,
        exp: Math.floor(Date.now() / 1000) + 12 * 60 * 60,
        sub: "mailto:soporte@lgtask.pages.dev"
      })
    )
  );
  const unsignedToken = `${header}.${payload}`;
  const privateKey = await crypto.subtle.importKey(
    "jwk",
    keys.privateJwk,
    { name: "ECDSA", namedCurve: "P-256" },
    false,
    ["sign"]
  );
  const signature = new Uint8Array(
    await crypto.subtle.sign({ name: "ECDSA", hash: "SHA-256" }, privateKey, new TextEncoder().encode(unsignedToken))
  );
  const token = `${unsignedToken}.${toBase64Url(signature)}`;
  return fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `vapid t=${token}, k=${keys.publicKey}`,
      "Crypto-Key": `p256ecdsa=${keys.publicKey}`,
      TTL: "86400",
      Urgency: "normal"
    }
  });
}

function mimeTypeForExtension(extension) {
  const types = {
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    webp: "image/webp",
    pdf: "application/pdf",
    doc: "application/msword",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xls: "application/vnd.ms-excel",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ppt: "application/vnd.ms-powerpoint",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    csv: "text/csv",
    txt: "text/plain"
  };
  return types[extension] || "application/octet-stream";
}

async function createUser(request, db, actor) {
  if (actor.role !== "Coordinador") return json({ ok: false, message: "Permiso insuficiente." }, 403);
  const body = await readJson(request);
  const email = clean(body.email).toLowerCase();
  const password = String(body.password || "");
  if (clean(body.name).length < 3 || !validEmail(email) || password.length < 6) {
    return json({ ok: false, message: "Completa los datos del trainer." }, 400);
  }
  const existing = await db.prepare("SELECT id FROM users WHERE email = ?").bind(email).first();
  if (existing) return json({ ok: false, message: "Ese correo ya existe." }, 409);
  const passwordData = await hashPassword(password);
  const createdAt = Date.now();
  const userId = crypto.randomUUID();
  await db
    .prepare("INSERT INTO users (id, name, email, zone, role, status, password_hash, password_salt, created_at) VALUES (?, ?, ?, ?, 'Trainer', 'Activo', ?, ?, ?)")
    .bind(userId, clean(body.name), email, clean(body.zone), passwordData.hash, passwordData.salt, createdAt)
    .run();
  return json({ ok: true, user: publicUser({ id: userId, name: clean(body.name), email, zone: clean(body.zone), role: "Trainer", status: "Activo", createdAt }) }, 201);
}

async function resetPassword(request, db, actor) {
  if (actor.role !== "Coordinador") return json({ ok: false, message: "Permiso insuficiente." }, 403);
  const body = await readJson(request);
  const user = await db.prepare("SELECT id, name, email FROM users WHERE id = ?").bind(clean(body.userId)).first();
  if (!user) return json({ ok: false, message: "Usuario no encontrado." }, 404);
  const tempPassword = `LG${randomToken(8).replace(/[-_]/g, "A").slice(0, 8)}`;
  const passwordData = await hashPassword(tempPassword);
  await db
    .prepare("UPDATE users SET password_hash = ?, password_salt = ? WHERE id = ?")
    .bind(passwordData.hash, passwordData.salt, user.id)
    .run();
  await db.prepare("DELETE FROM sessions WHERE user_id = ?").bind(user.id).run();

  if (body.requestId) {
    const data = await loadData(db);
    const requestItem = data.passwordRecoveryRequests.find((item) => item.id === body.requestId);
    if (requestItem) {
      requestItem.status = "Clave temporal generada";
      requestItem.resolvedAt = Date.now();
      requestItem.resolvedById = actor.id;
      await saveData(db, data);
    }
  }
  return json({ ok: true, tempPassword, user: publicUser(user) });
}

function timeToMinutes(value) {
  const match = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(String(value || ""));
  return match ? Number(match[1]) * 60 + Number(match[2]) : -1;
}

function validWorkSettings(settings) {
  const start = timeToMinutes(settings?.breakStart);
  const end = timeToMinutes(settings?.breakEnd);
  return (
    start >= timeToMinutes(WORKDAY_START) &&
    end <= timeToMinutes(WEEKDAY_END) &&
    end - start >= 15 &&
    end - start <= 180
  );
}

function normalizeWorkSettings(settings) {
  return validWorkSettings(settings)
    ? { breakStart: clean(settings.breakStart), breakEnd: clean(settings.breakEnd) }
    : { breakStart: BREAK_START, breakEnd: BREAK_END };
}

function normalizeBreakSettingsByUser(settingsByUser) {
  if (!settingsByUser || typeof settingsByUser !== "object" || Array.isArray(settingsByUser)) return {};
  return Object.fromEntries(
    Object.entries(settingsByUser)
      .filter(([userId, settings]) => clean(userId) && validWorkSettings(settings))
      .map(([userId, settings]) => [clean(userId), normalizeWorkSettings(settings)])
  );
}

function breakSettingsForUser(data, userId) {
  const personal = data.breakSettingsByUser?.[clean(userId)];
  return validWorkSettings(personal) ? normalizeWorkSettings(personal) : normalizeWorkSettings(data.workSettings);
}

function normalizeTaskCategory(category) {
  if (category === "PDV") return "PDP";
  if (category === "Producto") return "ULG";
  return ["PDP", "Entrenamiento", "ULG", "Reporte", "Coordinacion"].includes(category) ? category : "PDP";
}

function validWorkSchedule(dateValue, startTime, endTime, workSettings = null) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateValue || ""))) return false;
  const day = new Date(`${dateValue}T12:00:00Z`).getUTCDay();
  if (day === 0) return false;
  const start = timeToMinutes(startTime);
  const end = timeToMinutes(endTime);
  const workEnd = timeToMinutes(day === 6 ? SATURDAY_END : WEEKDAY_END);
  if (start < timeToMinutes(WORKDAY_START) || end > workEnd || end <= start) return false;
  const settings = normalizeWorkSettings(workSettings);
  if (day !== 6 && start < timeToMinutes(settings.breakEnd) && end > timeToMinutes(settings.breakStart)) return false;
  return true;
}

function hasTaskConflict(tasks, ownerId, dateValue, startTime, endTime, excludeTaskId = "") {
  const start = timeToMinutes(startTime);
  const end = timeToMinutes(endTime);
  return tasks.some((task) => {
    if (
      task.id === excludeTaskId ||
      task.ownerId !== ownerId ||
      task.dueDate !== dateValue ||
      task.status === "No disponible"
    ) {
      return false;
    }
    return start < timeToMinutes(task.endTime || "09:30") && end > timeToMinutes(task.startTime || WORKDAY_START);
  });
}

async function loadData(db) {
  const row = await db.prepare("SELECT data FROM app_data WHERE id = 1").first();
  try {
    const parsed = JSON.parse(row?.data || "{}");
    return {
      ...structuredClone(EMPTY_DATA),
      ...parsed,
      workSettings: normalizeWorkSettings(parsed.workSettings),
      breakSettingsByUser: normalizeBreakSettingsByUser(parsed.breakSettingsByUser),
      tasks: (parsed.tasks || []).map((task) => ({ ...task, category: normalizeTaskCategory(task.category) })),
      deletedTasks: (parsed.deletedTasks || []).map((task) => ({ ...task, category: normalizeTaskCategory(task.category) }))
    };
  } catch {
    return structuredClone(EMPTY_DATA);
  }
}

async function saveData(db, data) {
  const payload = {
    version: 9,
    workSettings: normalizeWorkSettings(data.workSettings),
    breakSettingsByUser: normalizeBreakSettingsByUser(data.breakSettingsByUser),
    registrationRequests: data.registrationRequests || [],
    passwordRecoveryRequests: data.passwordRecoveryRequests || [],
    tasks: data.tasks || [],
    deletedTasks: data.deletedTasks || [],
    announcements: data.announcements || [],
    supportRequests: data.supportRequests || [],
    dailyMotivations: data.dailyMotivations || []
  };
  await db.prepare("UPDATE app_data SET data = ?, updated_at = ? WHERE id = 1").bind(JSON.stringify(payload), Date.now()).run();
}

function publicUser(row) {
  return {
    id: row.id,
    name: row.name,
    email: row.email,
    zone: row.zone || "",
    role: row.role,
    status: row.status || "Activo",
    createdAt: Number(row.createdAt || row.created_at || Date.now())
  };
}

async function readJson(request, maxLength = 100000) {
  const text = await request.text();
  if (text.length > maxLength) throw new Error("Solicitud demasiado grande");
  return text ? JSON.parse(text) : {};
}

function sameOrigin(request) {
  const origin = request.headers.get("Origin");
  return !origin || origin === new URL(request.url).origin;
}

function validEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function clean(value) {
  return String(value || "").trim();
}

function readCookie(request, name) {
  const cookies = request.headers.get("Cookie") || "";
  for (const part of cookies.split(";")) {
    const [key, ...value] = part.trim().split("=");
    if (key === name) return decodeURIComponent(value.join("="));
  }
  return "";
}

function cookieValue(value, maxAge) {
  return `${SESSION_COOKIE}=${encodeURIComponent(value)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${maxAge}`;
}

async function hashPassword(password, salt = randomToken(16)) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: fromBase64Url(salt), iterations: 100000 },
    key,
    256
  );
  return { salt, hash: toBase64Url(new Uint8Array(bits)) };
}

async function verifyPassword(password, salt, expectedHash) {
  const result = await hashPassword(password, salt);
  return timingSafeEqual(result.hash, expectedHash);
}

function timingSafeEqual(left, right) {
  if (left.length !== right.length) return false;
  let value = 0;
  for (let i = 0; i < left.length; i += 1) value |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return value === 0;
}

function randomToken(length) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return toBase64Url(bytes);
}

async function sha256Hex(value) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function toBase64Url(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/g, "");
}

function fromBase64Url(value) {
  const padded = value.replaceAll("-", "+").replaceAll("_", "/") + "===".slice((value.length + 3) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

function json(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff"
    }
  });
}

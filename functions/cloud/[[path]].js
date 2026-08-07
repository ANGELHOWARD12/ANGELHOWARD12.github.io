const SESSION_COOKIE = "lg_session";
const SESSION_DAYS = 14;
const COORDINATOR_CODE_HASH = "e62163b1947feab8e4db70a99cffd5fb9c9f66d5e8901a4fb9775180ea780b71";
const MICROSOFT_STORAGE_PROVIDER = "onedrive";
const R2_STORAGE_PROVIDER = "r2";
const MAINTENANCE_INTERVAL_MS = 12 * 60 * 60 * 1000;
const ABANDONED_UPLOAD_TTL_MS = 10 * 60 * 1000;
const DELIVERED_NOTIFICATION_TTL_MS = 14 * 24 * 60 * 60 * 1000;
const UNDELIVERED_NOTIFICATION_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const MAINTENANCE_SETTING_KEY = "storage_maintenance_r2_direct_v4";
const OBSERVER_ACCESS_LEVEL = "observer";
const OBSERVER_EMAIL = "giuliana.parra@lgtask.local";

let microsoftTokenCache = { token: "", expiresAt: 0 };
let schemaReady = false;

const EMPTY_DATA = {
  version: 23,
  workSettings: {
    breakStart: "12:30",
    breakEnd: "14:00"
  },
  breakSettingsByUser: {},
  breakSettingsByUserDate: {},
  workScheduleByUserDate: {},
  overtimeRequests: [],
  registrationRequests: [],
  passwordRecoveryRequests: [],
  tasks: [],
  deletedTasks: [],
  announcements: [],
  supportRequests: [],
  dailyMotivations: [],
  lgUpdates: [],
  lateEvidenceUploadsEnabled: false,
  lateEvidencePolicyHistory: []
};

const WORKDAY_START = "08:30";
const WEEKDAY_END = "19:00";
const SATURDAY_END = "14:00";
const DANNY_SATURDAY_START = "09:00";
const DANNY_SATURDAY_END = "14:30";
const BREAK_START = "12:30";
const BREAK_END = "14:00";
const MAX_FILE_BASE64 = 1_800_000;
const MAX_FILE_TOTAL_BASE64 = 21_000_000;
const MAX_FILE_CHUNK_BASE64 = 900_000;
const MAX_FILE_BYTES = 15_000_000;
const ALLOWED_FILE_EXTENSIONS = new Set([
  "jpg",
  "jpeg",
  "png",
  "webp",
  "heic",
  "heif",
  "mp4",
  "mov",
  "webm",
  "m4v",
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
  "image/heic",
  "image/heif",
  "video/mp4",
  "video/quicktime",
  "video/webm",
  "video/x-m4v",
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
    scheduleMaintenance(env.DB, env, context);
    const url = new URL(request.url);
    const route = url.pathname.replace(/^\/(?:api|cloud)\/?/, "");

    if (request.method === "OPTIONS") return new Response(null, { status: 204 });
    if (!["GET", "HEAD"].includes(request.method) && !sameOrigin(request)) {
      return json({ ok: false, message: "Solicitud no permitida." }, 403);
    }

    if (route === "health" && request.method === "GET") return json({ ok: true, version: 23 });
    if (route === "auth/register" && request.method === "POST") return register(request, env.DB);
    if (route === "auth/login" && request.method === "POST") return login(request, env.DB);
    if (route === "auth/logout" && request.method === "POST") return logout(request, env.DB);
    if (route === "auth/recovery" && request.method === "POST") return requestRecovery(request, env.DB);

    const session = await authenticate(request, env.DB);
    if (!session) return json({ ok: false, message: "Sesion no valida." }, 401);
    if (
      isObserverUser(session.user) &&
      !["GET", "HEAD"].includes(request.method) &&
      !["notifications/subscribe", "notifications/unsubscribe"].includes(route)
    ) {
      return json({ ok: false, message: "La cuenta observadora es de solo lectura." }, 403);
    }

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
    if (route === "storage/status" && request.method === "GET") return storageStatus(env.DB, session.user, env);
    if (route === "evidence/upload/r2" && request.method === "POST") {
      return uploadEvidenceDirectToR2(request, env.DB, session.user, env);
    }
    if (route === "evidence/upload" && request.method === "POST") return uploadEvidence(request, env.DB, session.user, env);
    if (route === "evidence/upload/init" && request.method === "POST") return initEvidenceUpload(request, env.DB, session.user);
    if (route === "evidence/upload/chunk" && request.method === "POST") return uploadEvidenceChunk(request, env.DB, session.user);
    if (route === "evidence/upload/complete" && request.method === "POST") return completeEvidenceUpload(request, env.DB, session.user, env);
    const evidenceFileMatch = route.match(/^evidence\/([^/]+)\/(?:file|photo)$/);
    if (evidenceFileMatch && request.method === "GET") return evidenceFile(env.DB, session.user, evidenceFileMatch[1], env);
    if (route === "tasks/evidence" && request.method === "POST") return submitTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/create" && request.method === "POST") return createTask(request, env.DB, session.user, context);
    if (route === "tasks/evidence-authorize" && request.method === "POST") return authorizeLateTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/review" && request.method === "POST") return reviewTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/delete" && request.method === "POST") return deleteTaskAndArchive(request, env.DB, session.user, context, env);
    if (route === "schedule/overtime" && request.method === "POST") return saveOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "schedule/overtime-request" && request.method === "POST") return requestOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "schedule/overtime-review" && request.method === "POST") return reviewOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "state" && request.method === "GET") return getState(env.DB, session.user, context);
    if (route === "state" && request.method === "PUT") return putState(request, env.DB, session.user, context);
    if (route === "settings/late-evidence" && request.method === "POST") {
      return setLateEvidencePolicy(request, env.DB, session.user, context);
    }
    if (route === "info/updates" && request.method === "POST") return saveInfoUpdate(request, env.DB, session.user);
    if (route === "admin/users" && request.method === "POST") return createUser(request, env.DB, session.user);
    if (route === "admin/reset-password" && request.method === "POST") return resetPassword(request, env.DB, session.user);

    return json({ ok: false, message: "Ruta no encontrada." }, 404);
  } catch (error) {
    console.error("LG Task API", error);
    return json({ ok: false, message: "No se pudo completar la operacion en la nube." }, 500);
  }
}

async function ensureSchema(db) {
  if (schemaReady) return;
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
    db.prepare(`CREATE TABLE IF NOT EXISTS evidence_storage (
      file_id TEXT PRIMARY KEY,
      provider TEXT NOT NULL,
      drive_id TEXT NOT NULL,
      drive_item_id TEXT NOT NULL,
      parent_path TEXT NOT NULL DEFAULT '',
      size_bytes INTEGER NOT NULL DEFAULT 0,
      sha256 TEXT NOT NULL DEFAULT '',
      created_at INTEGER NOT NULL,
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
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_storage_item ON evidence_storage(drive_item_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_notification_subscriptions_user ON notification_subscriptions(user_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_user_notifications_pending ON user_notifications(user_id, delivered_at)")
  ]);
  const userColumns = await db.prepare("PRAGMA table_info(users)").all();
  if (!(userColumns.results || []).some((column) => column.name === "access_level")) {
    try {
      await db.prepare("ALTER TABLE users ADD COLUMN access_level TEXT NOT NULL DEFAULT ''").run();
    } catch (error) {
      if (!String(error?.message || error).toLowerCase().includes("duplicate column")) throw error;
    }
  }
  await db
    .prepare("UPDATE users SET access_level = ?, zone = 'Administracion general' WHERE email = ?")
    .bind(OBSERVER_ACCESS_LEVEL, OBSERVER_EMAIL)
    .run();
  schemaReady = true;
}

function scheduleMaintenance(db, env, context) {
  context.waitUntil(
    runMaintenance(db, env).catch((error) => {
      console.error("LGTASK maintenance", error);
    })
  );
}

async function runMaintenance(db, env) {
  const now = Date.now();
  const claim = await db
    .prepare(
      `INSERT INTO app_settings (key, value, updated_at) VALUES (?, 'running', ?)
       ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
       WHERE app_settings.updated_at < ?`
    )
    .bind(MAINTENANCE_SETTING_KEY, now, now - MAINTENANCE_INTERVAL_MS)
    .run();
  if (Number(claim?.meta?.changes || 0) === 0) return;

  const abandonedCutoff = now - ABANDONED_UPLOAD_TTL_MS;
  const notificationCutoff = now - DELIVERED_NOTIFICATION_TTL_MS;
  const undeliveredNotificationCutoff = now - UNDELIVERED_NOTIFICATION_TTL_MS;
  const unreferencedFileIds = `
    SELECT ef.id FROM evidence_files ef
    WHERE ef.created_at < ?
      AND NOT EXISTS (
        SELECT 1
        FROM app_data a,
             json_each(a.data, '$.tasks') task,
             json_each(task.value, '$.evidence') evidence,
             json_each(evidence.value, '$.files') file
        WHERE a.id = 1 AND json_extract(file.value, '$.id') = ef.id
      )`;

  const abandonedStoredItems = await db
    .prepare(
      `SELECT storage.provider, storage.drive_id, storage.drive_item_id
       FROM evidence_storage storage
       INNER JOIN evidence_files files ON files.id = storage.file_id
       WHERE files.id IN (${unreferencedFileIds})`
    )
    .bind(abandonedCutoff)
    .all();
  if (abandonedStoredItems.results?.length) {
    await deleteExternalEvidenceItems(env, abandonedStoredItems.results);
  }

  await db.batch([
    db.prepare("DELETE FROM evidence_file_chunks WHERE file_id IN (SELECT file_id FROM evidence_storage)"),
    db.prepare("DELETE FROM evidence_file_chunks WHERE file_id NOT IN (SELECT id FROM evidence_files)"),
    db.prepare(`DELETE FROM evidence_file_chunks WHERE file_id IN (${unreferencedFileIds})`).bind(abandonedCutoff),
    db.prepare(`DELETE FROM evidence_storage WHERE file_id IN (${unreferencedFileIds})`).bind(abandonedCutoff),
    db.prepare(`DELETE FROM evidence_files WHERE id IN (${unreferencedFileIds})`).bind(abandonedCutoff),
    db.prepare("DELETE FROM sessions WHERE expires_at < ?").bind(now),
    db.prepare("DELETE FROM user_notifications WHERE delivered_at > 0 AND created_at < ?").bind(notificationCutoff),
    db.prepare("DELETE FROM user_notifications WHERE delivered_at = 0 AND created_at < ?").bind(undeliveredNotificationCutoff)
  ]);

  const migration = await migrateOneLegacyEvidence(db, env);
  if (migration.migrated || migration.failed) {
    await db
      .prepare("UPDATE app_settings SET value = ?, updated_at = 0 WHERE key = ?")
      .bind(migration.migrated ? "migrating" : "migration-retry", MAINTENANCE_SETTING_KEY)
      .run();
    return;
  }

  await db
    .prepare("UPDATE app_settings SET value = 'complete', updated_at = ? WHERE key = ?")
    .bind(Date.now(), MAINTENANCE_SETTING_KEY)
    .run();
}

async function migrateOneLegacyEvidence(db, env) {
  if (!r2StorageEnabled(env)) return { migrated: false, failed: false };
  const candidate = await db
    .prepare(
      `SELECT files.*
       FROM evidence_files files
       WHERE NOT EXISTS (SELECT 1 FROM evidence_storage storage WHERE storage.file_id = files.id)
         AND (LENGTH(files.photo_base64) > 0 OR EXISTS (
           SELECT 1 FROM evidence_file_chunks chunks WHERE chunks.file_id = files.id
         ))
         AND EXISTS (
           SELECT 1
           FROM app_data data,
                json_each(data.data, '$.tasks') task,
                json_each(task.value, '$.evidence') evidence,
                json_each(evidence.value, '$.files') file
           WHERE data.id = 1 AND json_extract(file.value, '$.id') = files.id
         )
         AND NOT EXISTS (
           SELECT 1 FROM app_settings setting WHERE setting.key = 'storage_migration_failed:' || files.id
         )
       ORDER BY files.created_at ASC
       LIMIT 1`
    )
    .first();
  if (!candidate) return { migrated: false, failed: false };

  try {
    const data = await loadData(db);
    const task = data.tasks.find((item) => clean(item.id) === clean(candidate.task_id));
    if (!task) throw new Error("Referenced task not found");
    const file = { id: candidate.id, fileName: candidate.file_name, mimeType: candidate.mime_type };
    if (candidate.photo_base64) {
      await uploadEvidenceToR2(env, db, task, file, candidate.photo_base64);
    } else {
      await uploadChunkedEvidenceToR2(env, db, task, file);
    }
    await db.batch([
      db.prepare("UPDATE evidence_files SET photo_base64 = '' WHERE id = ?").bind(candidate.id),
      db.prepare("DELETE FROM evidence_file_chunks WHERE file_id = ?").bind(candidate.id)
    ]);
    return { migrated: true, failed: false };
  } catch (error) {
    console.error("Legacy evidence migration", candidate.id, error);
    await db
      .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
      .bind(`storage_migration_failed:${candidate.id}`, clean(error?.message || error).slice(0, 300), Date.now())
      .run();
    return { migrated: false, failed: true };
  }
}

async function storageStatus(db, user, env) {
  if (user.role !== "Coordinador" && !isObserverUser(user)) {
    return json({ ok: false, message: "Solo administracion puede consultar el almacenamiento." }, 403);
  }

  const counts = await db
    .prepare(
      `SELECT
         (SELECT COUNT(*) FROM evidence_files) AS evidence_files,
         (SELECT COUNT(*) FROM evidence_storage) AS stored_files,
         (SELECT COUNT(*) FROM evidence_file_chunks) AS temporary_chunks,
         (SELECT COALESCE(SUM(LENGTH(photo_base64)), 0) FROM evidence_files) AS legacy_base64_bytes,
         (SELECT COALESCE(SUM(LENGTH(chunk_base64)), 0) FROM evidence_file_chunks) AS temporary_base64_bytes,
         (SELECT COALESCE(LENGTH(data), 0) FROM app_data WHERE id = 1) AS app_data_bytes,
         (SELECT COUNT(*) FROM sessions WHERE expires_at < ?) AS expired_sessions,
         (SELECT COUNT(*) FROM user_notifications WHERE delivered_at > 0 AND created_at < ?) AS old_notifications`
    )
    .bind(Date.now(), Date.now() - DELIVERED_NOTIFICATION_TTL_MS)
    .first();
  const chunkFiles = await db
    .prepare(
      `SELECT files.id, files.created_at, COUNT(chunks.chunk_index) AS chunks,
              COALESCE(SUM(LENGTH(chunks.chunk_base64)), 0) AS bytes
       FROM evidence_files files
       INNER JOIN evidence_file_chunks chunks ON chunks.file_id = files.id
       GROUP BY files.id, files.created_at`
    )
    .all();
  const appData = await loadData(db);
  const referencedFileIds = new Set(
    (appData.tasks || []).flatMap((task) =>
      (Array.isArray(task.evidence) ? task.evidence : []).flatMap((entry) =>
        (Array.isArray(entry.files) ? entry.files : []).map((file) => clean(file.id)).filter(Boolean)
      )
    )
  );
  const temporaryFiles = chunkFiles.results || [];
  const unreferencedTemporaryFiles = temporaryFiles.filter((file) => !referencedFileIds.has(clean(file.id)));

  let pageCount = 0;
  let pageSize = 0;
  let freePages = 0;
  try {
    pageCount = Number((await db.prepare("PRAGMA page_count").first())?.page_count || 0);
    pageSize = Number((await db.prepare("PRAGMA page_size").first())?.page_size || 0);
    freePages = Number((await db.prepare("PRAGMA freelist_count").first())?.freelist_count || 0);
  } catch (error) {
    console.error("D1 storage diagnostics", error);
  }

  let r2Objects = 0;
  let r2Bytes = 0;
  let r2Complete = true;
  if (r2StorageEnabled(env)) {
    let cursor;
    let pages = 0;
    do {
      const listed = await env.EVIDENCE_BUCKET.list({ limit: 1000, cursor });
      r2Objects += listed.objects.length;
      r2Bytes += listed.objects.reduce((total, item) => total + Number(item.size || 0), 0);
      cursor = listed.truncated ? listed.cursor : undefined;
      pages += 1;
      if (pages >= 100 && cursor) {
        r2Complete = false;
        break;
      }
    } while (cursor);
  }

  const maintenance = await db
    .prepare("SELECT value, updated_at FROM app_settings WHERE key = ?")
    .bind(MAINTENANCE_SETTING_KEY)
    .first();
  const failedMigrations = await db
    .prepare("SELECT COUNT(*) AS count FROM app_settings WHERE key LIKE 'storage_migration_failed:%'")
    .first();
  return json({
    ok: true,
    storage: {
      d1: {
        evidenceFiles: Number(counts?.evidence_files || 0),
        storedFiles: Number(counts?.stored_files || 0),
        temporaryChunks: Number(counts?.temporary_chunks || 0),
        temporaryFiles: temporaryFiles.length,
        referencedTemporaryFiles: temporaryFiles.length - unreferencedTemporaryFiles.length,
        unreferencedTemporaryFiles: unreferencedTemporaryFiles.length,
        unreferencedTemporaryBytes: unreferencedTemporaryFiles.reduce((total, file) => total + Number(file.bytes || 0), 0),
        legacyBase64Bytes: Number(counts?.legacy_base64_bytes || 0),
        temporaryBase64Bytes: Number(counts?.temporary_base64_bytes || 0),
        appDataBytes: Number(counts?.app_data_bytes || 0),
        approximatePayloadBytes:
          Number(counts?.legacy_base64_bytes || 0) +
          Number(counts?.temporary_base64_bytes || 0) +
          Number(counts?.app_data_bytes || 0),
        databaseBytes: pageCount && pageSize ? pageCount * pageSize : 0,
        reusableBytes: freePages && pageSize ? freePages * pageSize : 0,
        expiredSessions: Number(counts?.expired_sessions || 0),
        oldNotifications: Number(counts?.old_notifications || 0)
      },
      r2: {
        enabled: r2StorageEnabled(env),
        objects: r2Objects,
        bytes: r2Bytes,
        complete: r2Complete
      },
      maintenance: {
        state: clean(maintenance?.value) || "pending",
        updatedAt: Number(maintenance?.updated_at || 0),
        failedMigrations: Number(failedMigrations?.count || 0)
      }
    }
  });
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

async function setLateEvidencePolicy(request, db, user, context) {
  if (user.role !== "Coordinador") {
    return json({ ok: false, message: "Solo Pablo puede cambiar el permiso de sustentos anteriores." }, 403);
  }
  const body = await readJson(request);
  if (typeof body.enabled !== "boolean") {
    return json({ ok: false, message: "Selecciona si la carga de sustentos anteriores debe estar encendida o apagada." }, 400);
  }

  const data = await loadData(db);
  const enabled = body.enabled;
  const changedAt = Date.now();
  data.lateEvidenceUploadsEnabled = enabled;
  data.lateEvidencePolicyHistory = Array.isArray(data.lateEvidencePolicyHistory)
    ? data.lateEvidencePolicyHistory
    : [];
  data.lateEvidencePolicyHistory.unshift({
    enabled,
    byId: user.id,
    byName: user.name,
    at: changedAt
  });
  data.lateEvidencePolicyHistory = data.lateEvidencePolicyHistory.slice(0, 100);
  await saveData(db, data);

  const trainers = await db
    .prepare("SELECT id FROM users WHERE role = 'Trainer' AND status = 'Activo' AND access_level <> ?")
    .bind(OBSERVER_ACCESS_LEVEL)
    .all();
  const notifiedUsers = await queueNotifications(
    db,
    (trainers.results || []).map((trainer) => ({
      userId: trainer.id,
      title: enabled ? "Sustentos anteriores habilitados" : "Sustentos anteriores bloqueados",
      body: enabled
        ? "Pablo habilito temporalmente la carga de evidencias para tareas de dias anteriores."
        : "La carga de evidencias vuelve a estar disponible solamente para las tareas del dia.",
      url: "/?view=evidenceView",
      sourceKey: `late-evidence-policy:${enabled ? "on" : "off"}:${changedAt}:${trainer.id}`
    }))
  );
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function saveInfoUpdate(request, db, user) {
  const body = await readJson(request);
  const data = await loadData(db);
  const action = clean(body.action) || "upsert";
  const itemId = clean(body.id);

  if (action === "delete") {
    if (!itemId || !data.lgUpdates.some((item) => item.id === itemId)) {
      return json({ ok: false, message: "La publicacion de Info LG ya no existe." }, 404);
    }
    data.lgUpdates = data.lgUpdates.filter((item) => item.id !== itemId);
    await saveData(db, data);
    return stateResponse(db, user, data);
  }

  const line = ["HS", "TV", "AV"].includes(body.line) ? body.line : "";
  const type = ["Producto", "Lanzamiento", "Tecnologia", "Argumento de venta"].includes(body.type) ? body.type : "";
  const title = clean(body.title).slice(0, 140);
  const product = clean(body.product).slice(0, 120);
  const description = clean(body.description).slice(0, 2000);
  if (!line || !type || title.length < 3 || description.length < 5) {
    return json({ ok: false, message: "Completa linea, tipo, titulo e informacion antes de publicar." }, 400);
  }

  const now = Date.now();
  const existingIndex = itemId ? data.lgUpdates.findIndex((item) => item.id === itemId) : -1;
  const existing = existingIndex >= 0 ? data.lgUpdates[existingIndex] : null;
  const nextItem = {
    id: existing?.id || crypto.randomUUID(),
    line,
    type,
    title,
    product,
    description,
    createdById: existing?.createdById || user.id,
    createdByName: existing?.createdByName || user.name,
    createdAt: existing?.createdAt || now,
    updatedById: user.id,
    updatedByName: user.name,
    updatedAt: now
  };
  if (existingIndex >= 0) data.lgUpdates[existingIndex] = nextItem;
  else data.lgUpdates.unshift(nextItem);
  data.lgUpdates = normalizeLgUpdates(data.lgUpdates).slice(0, 500);
  await saveData(db, data);
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
  await queueWeeklyPlanningReminder(db, user);
  await queueEvidenceDeadlineReminder(db, user);
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

async function queueWeeklyPlanningReminder(db, user) {
  const today = dateIsoInLima(Date.now());
  const day = new Date(`${today}T12:00:00Z`).getUTCDay();
  if (day !== 5 && day !== 6) return;
  const currentMonday = addDaysIso(today, -(day - 1));
  const nextMonday = addDaysIso(currentMonday, 7);
  await queueNotifications(db, [
    {
      userId: user.id,
      title: `Prepara la Week ${isoWeekNumber(nextMonday)}`,
      body: "Completa el cronograma de lunes a sabado para la siguiente semana.",
      url: "/?view=tasksView",
      sourceKey: `weekly-planning:${clean(user.id)}:${nextMonday}`
    }
  ]);
}

async function queueEvidenceDeadlineReminder(db, user) {
  if (user.role !== "Trainer") return;
  const now = Date.now();
  const today = dateIsoInLima(now);
  const hour = Number(
    new Intl.DateTimeFormat("en", { timeZone: "America/Lima", hour: "2-digit", hourCycle: "h23" }).format(new Date(now))
  );
  if (hour < 17 || hour >= 19) return;
  const data = await loadData(db);
  const pending = (data.tasks || []).filter(
    (task) =>
      clean(task.ownerId) === user.id &&
      clean(task.dueDate) === today &&
      task.status !== "Cumplida" &&
      !(task.evidence || []).length
  );
  if (!pending.length) return;
  await queueNotifications(db, [
    {
      userId: user.id,
      title: "Sustentos pendientes de hoy",
      body: `${pending.length} tarea${pending.length === 1 ? "" : "s"} aun no ${pending.length === 1 ? "tiene" : "tienen"} evidencia. El plazo termina al finalizar el dia.`,
      url: "/?view=tasksView",
      sourceKey: `evidence-deadline:${clean(user.id)}:${today}`
    }
  ]);
}

function dateIsoInLima(value) {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en", {
      timeZone: "America/Lima",
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    })
      .formatToParts(new Date(value))
      .filter((part) => part.type !== "literal")
      .map((part) => [part.type, part.value])
  );
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function lateEvidenceAuthorizationActive(task, today = dateIsoInLima(Date.now())) {
  const authorization = task?.lateEvidenceAuthorization;
  return Boolean(authorization?.enabled && clean(authorization.validDate) === today);
}

function taskAllowsEvidenceUpload(task, today = dateIsoInLima(Date.now()), allowPreviousDays = false) {
  const dueDate = clean(task?.dueDate);
  return Boolean(
    task &&
      (dueDate === today ||
        (dueDate < today && Boolean(allowPreviousDays)) ||
        lateEvidenceAuthorizationActive(task, today))
  );
}

function evidenceUploadWindowError(task, today = dateIsoInLima(Date.now()), allowPreviousDays = false) {
  if (clean(task?.dueDate) < today) {
    if (allowPreviousDays) return "La carga de sustentos anteriores esta habilitada por Pablo.";
    return "El plazo para subir sustentos de esta tarea ha finalizado. Las evidencias deben registrarse el mismo dia de la ejecucion.";
  }
  if (clean(task?.dueDate) > today) return "El sustento solo puede registrarse el dia programado para la tarea.";
  return "La carga de sustentos no esta disponible para esta tarea.";
}

function addDaysIso(dateValue, days) {
  const date = new Date(`${dateValue}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function isoWeekNumber(dateValue) {
  const date = new Date(`${dateValue}T12:00:00Z`);
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
}

function historyIsAppendOnly(previousHistory, nextHistory) {
  const previous = Array.isArray(previousHistory) ? previousHistory : [];
  const next = Array.isArray(nextHistory) ? nextHistory : [];
  return next.length >= previous.length && previous.every((entry, index) => JSON.stringify(entry) === JSON.stringify(next[index]));
}

function microsoftStorageConfigured(env) {
  return Boolean(
    clean(env.MS_TENANT_ID) &&
      clean(env.MS_CLIENT_ID) &&
      clean(env.MS_CLIENT_SECRET) &&
      clean(env.MS_DRIVE_ID) &&
      clean(env.MS_ROOT_FOLDER_ID)
  );
}

function microsoftStorageEnabled(env) {
  return microsoftStorageConfigured(env) && clean(env.MS_STORAGE_ENABLED).toLowerCase() === "true";
}

async function microsoftAccessToken(env, forceRefresh = false) {
  if (!microsoftStorageConfigured(env)) throw new Error("Microsoft storage is not configured");
  if (!forceRefresh && microsoftTokenCache.token && microsoftTokenCache.expiresAt > Date.now() + 60_000) {
    return microsoftTokenCache.token;
  }
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(clean(env.MS_TENANT_ID))}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clean(env.MS_CLIENT_ID),
        client_secret: String(env.MS_CLIENT_SECRET || ""),
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      })
    }
  );
  if (!response.ok) throw new Error(`Microsoft token request failed (${response.status})`);
  const payload = await response.json();
  if (!payload.access_token) throw new Error("Microsoft token response is incomplete");
  microsoftTokenCache = {
    token: payload.access_token,
    expiresAt: Date.now() + Math.max(60, Number(payload.expires_in || 3600)) * 1000
  };
  return microsoftTokenCache.token;
}

async function graphFetch(env, path, init = {}, retry = true) {
  const headers = new Headers(init.headers || {});
  headers.set("Authorization", `Bearer ${await microsoftAccessToken(env)}`);
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, { ...init, headers });
  if (response.status === 401 && retry) {
    microsoftTokenCache = { token: "", expiresAt: 0 };
    return graphFetch(env, path, init, false);
  }
  return response;
}

function oneDriveSafeName(value, fallback, maxLength = 90) {
  return (
    clean(value)
      .replace(/[\\/:*?"<>|#%]/g, "-")
      .replace(/[. ]+$/g, "")
      .replace(/\s+/g, " ")
      .slice(0, maxLength) || fallback
  );
}

async function graphResponseError(response, operation) {
  let code = "unknown";
  try {
    const payload = await response.json();
    code = clean(payload?.error?.code) || code;
  } catch {
    // Graph occasionally returns an empty response for infrastructure errors.
  }
  return new Error(`${operation} failed (${response.status}, ${code})`);
}

async function ensureOneDriveFolder(env, parentId, folderName) {
  const safeName = oneDriveSafeName(folderName, "Sin nombre", 80);
  const driveId = encodeURIComponent(clean(env.MS_DRIVE_ID));
  const parent = encodeURIComponent(clean(parentId));
  const itemPath = `/drives/${driveId}/items/${parent}:/${encodeURIComponent(safeName)}`;
  let response = await graphFetch(env, itemPath);
  if (response.ok) return response.json();
  if (response.status !== 404) throw await graphResponseError(response, "OneDrive folder lookup");

  response = await graphFetch(env, `/drives/${driveId}/items/${parent}/children`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: safeName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail"
    })
  });
  if (response.ok) return response.json();
  if (response.status === 409) {
    const existing = await graphFetch(env, itemPath);
    if (existing.ok) return existing.json();
    throw await graphResponseError(existing, "OneDrive folder lookup after conflict");
  }
  throw await graphResponseError(response, "OneDrive folder creation");
}

function decodeEvidenceBase64(fileBase64) {
  const binary = atob(fileBase64);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

async function sha256Bytes(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function r2StorageEnabled(env) {
  return Boolean(env.EVIDENCE_BUCKET);
}

function externalStorageProvider(env) {
  if (r2StorageEnabled(env)) return R2_STORAGE_PROVIDER;
  if (microsoftStorageEnabled(env)) return MICROSOFT_STORAGE_PROVIDER;
  return "";
}

async function r2EvidenceDescriptor(db, task, file) {
  const owner = await db.prepare("SELECT name FROM users WHERE id = ?").bind(clean(task.ownerId)).first();
  const trainerName = oneDriveSafeName(owner?.name, "Trainer");
  const dueDate = oneDriveSafeName(task.dueDate, dateIsoInLima(Date.now()), 20);
  const taskName = oneDriveSafeName(task.title, "Tarea");
  const storedName = oneDriveSafeName(`${clean(file.id).slice(0, 8)} - ${file.fileName}`, "sustento", 120);
  const parentPath = `LGTASK Sustentos / ${trainerName} / ${dueDate} / ${taskName}`;
  const objectKey = `sustentos/${trainerName}/${dueDate}/${taskName}/${storedName}`;
  return { objectKey, parentPath };
}

async function uploadEvidenceToR2(env, db, task, file, fileBase64) {
  if (!r2StorageEnabled(env)) throw new Error("R2 storage is not configured");
  const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
  const bytes = decodeEvidenceBase64(fileBase64);

  await env.EVIDENCE_BUCKET.put(objectKey, bytes, {
    httpMetadata: { contentType: file.mimeType },
    customMetadata: { fileId: clean(file.id), taskId: clean(task.id), ownerId: clean(task.ownerId) }
  });
  try {
    await db
      .prepare(
        `INSERT OR REPLACE INTO evidence_storage
         (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      )
      .bind(
        file.id,
        R2_STORAGE_PROVIDER,
        "EVIDENCE_BUCKET",
        objectKey,
        parentPath,
        bytes.byteLength,
        await sha256Bytes(bytes),
        Date.now()
      )
      .run();
  } catch (error) {
    await env.EVIDENCE_BUCKET.delete(objectKey);
    throw error;
  }
  return objectKey;
}

async function uploadChunkedEvidenceToR2(env, db, task, file) {
  const chunks = await db
    .prepare(
      "SELECT chunk_index, chunk_base64 FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC"
    )
    .bind(file.id)
    .all();
  const chunkRows = chunks.results || [];
  if (!chunkRows.length || chunkRows.some((chunk, index) => Number(chunk.chunk_index) !== index)) {
    throw new Error("Incomplete legacy chunks");
  }
  const fileBase64 = chunkRows.map((chunk) => String(chunk.chunk_base64 || "")).join("");
  if (!fileBase64 || fileBase64.length > MAX_FILE_TOTAL_BASE64) throw new Error("Invalid legacy file size");
  return uploadEvidenceToR2(env, db, task, file, fileBase64);
}

async function uploadEvidenceDirectToR2(request, db, user, env) {
  if (!r2StorageEnabled(env)) {
    return json({ ok: false, message: "El almacenamiento R2 no esta configurado." }, 503);
  }

  const url = new URL(request.url);
  const taskId = clean(url.searchParams.get("taskId"));
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (user.role !== "Coordinador" && clean(task.ownerId) !== clean(user.id)) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  const today = dateIsoInLima(Date.now());
  if (!taskAllowsEvidenceUpload(task, today, data.lateEvidenceUploadsEnabled)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, today, data.lateEvidenceUploadsEnabled) }, 409);
  }

  const fileName = clean(url.searchParams.get("fileName")).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "archivo-sustento";
  const extension = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
  const mimeType = clean(request.headers.get("Content-Type")).split(";", 1)[0].toLowerCase() || mimeTypeForExtension(extension);
  const declaredSize = Number(url.searchParams.get("size") || request.headers.get("Content-Length") || 0);
  const contentLength = Number(request.headers.get("Content-Length") || 0);
  if (
    !request.body ||
    !ALLOWED_FILE_EXTENSIONS.has(extension) ||
    !ALLOWED_FILE_MIME_TYPES.has(mimeType) ||
    !Number.isInteger(declaredSize) ||
    declaredSize < 1 ||
    declaredSize > MAX_FILE_BYTES ||
    (contentLength && contentLength !== declaredSize)
  ) {
    return json({ ok: false, message: "El archivo no tiene un formato o tamano permitido." }, 400);
  }

  const fileId = crypto.randomUUID();
  const createdAt = Date.now();
  const file = { id: fileId, fileName, mimeType };
  const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
  let storedObject;
  try {
    storedObject = await env.EVIDENCE_BUCKET.put(objectKey, request.body, {
      httpMetadata: { contentType: mimeType },
      customMetadata: { fileId, taskId: clean(task.id), ownerId: clean(task.ownerId) }
    });
  } catch (error) {
    console.error("Direct R2 evidence upload", error);
    return json({ ok: false, message: "R2 no pudo recibir el archivo. Intenta nuevamente." }, 503);
  }

  const storedSize = Number(storedObject?.size || 0);
  if (storedSize < 1 || storedSize > MAX_FILE_BYTES || storedSize !== declaredSize) {
    await env.EVIDENCE_BUCKET.delete(objectKey).catch(() => {});
    return json({ ok: false, message: "La carga quedo incompleta y fue descartada. Intenta nuevamente." }, 409);
  }

  try {
    await db.batch([
      db
        .prepare(
          "INSERT INTO evidence_files (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at) VALUES (?, ?, ?, ?, ?, ?, '', ?)"
        )
        .bind(fileId, task.id, task.ownerId, user.id, fileName, mimeType, createdAt),
      db
        .prepare(
          `INSERT INTO evidence_storage
           (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
           VALUES (?, ?, ?, ?, ?, ?, '', ?)`
        )
        .bind(fileId, R2_STORAGE_PROVIDER, "EVIDENCE_BUCKET", objectKey, parentPath, storedSize, createdAt)
    ]);
  } catch (error) {
    await env.EVIDENCE_BUCKET.delete(objectKey).catch(() => {});
    console.error("Direct R2 evidence metadata", error);
    return json({ ok: false, message: "El archivo llego a R2, pero no se pudo registrar. Intenta nuevamente." }, 503);
  }

  return json({
    ok: true,
    file: { id: fileId, name: fileName, mimeType, createdAt, url: `/cloud/evidence/${fileId}/file` }
  }, 201);
}

async function r2FileResponse(env, storage, row) {
  if (!r2StorageEnabled(env)) {
    return json({ ok: false, message: "El almacenamiento R2 no esta configurado." }, 503);
  }
  const object = await env.EVIDENCE_BUCKET.get(clean(storage.drive_item_id));
  if (!object) return json({ ok: false, message: "Archivo no encontrado en R2." }, 404);
  const safeName = String(row.file_name || "sustento.jpg").replace(/["\r\n]/g, "-");
  const disposition =
    String(row.mime_type).startsWith("image/") ||
    String(row.mime_type).startsWith("video/") ||
    row.mime_type === "application/pdf"
      ? "inline"
      : "attachment";
  const headers = new Headers({
    "Content-Type": row.mime_type,
    "Content-Disposition": `${disposition}; filename="${safeName}"`,
    "Cache-Control": "private, max-age=300",
    "Content-Length": String(object.size),
    "X-Content-Type-Options": "nosniff"
  });
  if (object.httpEtag) headers.set("ETag", object.httpEtag);
  return new Response(object.body, { status: 200, headers });
}

async function uploadEvidenceToOneDrive(env, db, task, file, fileBase64) {
  const owner = await db.prepare("SELECT name FROM users WHERE id = ?").bind(clean(task.ownerId)).first();
  const trainerName = oneDriveSafeName(owner?.name, "Trainer");
  const dueDate = oneDriveSafeName(task.dueDate, dateIsoInLima(Date.now()), 20);
  const taskName = oneDriveSafeName(task.title, "Tarea");
  const trainerFolder = await ensureOneDriveFolder(env, clean(env.MS_ROOT_FOLDER_ID), trainerName);
  const dateFolder = await ensureOneDriveFolder(env, trainerFolder.id, dueDate);
  const taskFolder = await ensureOneDriveFolder(env, dateFolder.id, taskName);
  const storedName = oneDriveSafeName(`${clean(file.id).slice(0, 8)} - ${file.fileName}`, "sustento", 120);
  const bytes = decodeEvidenceBase64(fileBase64);
  const driveId = encodeURIComponent(clean(env.MS_DRIVE_ID));
  const uploadPath = `/drives/${driveId}/items/${encodeURIComponent(taskFolder.id)}:/${encodeURIComponent(storedName)}:/content`;
  const response = await graphFetch(env, uploadPath, {
    method: "PUT",
    headers: { "Content-Type": file.mimeType },
    body: bytes
  });
  if (!response.ok) throw await graphResponseError(response, "OneDrive file upload");
  const item = await response.json();
  const parentPath = `LGTASK Sustentos / ${trainerName} / ${dueDate} / ${taskName}`;
  await db
    .prepare(
      `INSERT OR REPLACE INTO evidence_storage
       (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .bind(
      file.id,
      MICROSOFT_STORAGE_PROVIDER,
      clean(env.MS_DRIVE_ID),
      clean(item.id),
      parentPath,
      bytes.byteLength,
      await sha256Bytes(bytes),
      Date.now()
    )
    .run();
  return item;
}

async function oneDriveFileResponse(env, storage, row) {
  if (!microsoftStorageConfigured(env)) {
    return json({ ok: false, message: "El almacenamiento de OneDrive no esta configurado." }, 503);
  }
  const path = `/drives/${encodeURIComponent(storage.drive_id)}/items/${encodeURIComponent(storage.drive_item_id)}/content`;
  let response = await graphFetch(env, path, { redirect: "manual" });
  if (response.status >= 300 && response.status < 400 && response.headers.get("Location")) {
    response = await fetch(response.headers.get("Location"), { redirect: "follow" });
  }
  if (!response.ok) return json({ ok: false, message: "No se pudo abrir el archivo de OneDrive." }, response.status === 404 ? 404 : 503);
  const safeName = String(row.file_name || "sustento.jpg").replace(/["\r\n]/g, "-");
  const disposition =
    String(row.mime_type).startsWith("image/") ||
    String(row.mime_type).startsWith("video/") ||
    row.mime_type === "application/pdf"
      ? "inline"
      : "attachment";
  const headers = new Headers({
    "Content-Type": row.mime_type,
    "Content-Disposition": `${disposition}; filename="${safeName}"`,
    "Cache-Control": "private, max-age=300",
    "X-Content-Type-Options": "nosniff"
  });
  const contentLength = response.headers.get("Content-Length") || String(storage.size_bytes || "");
  if (contentLength) headers.set("Content-Length", contentLength);
  return new Response(response.body, { status: 200, headers });
}

async function uploadEvidence(request, db, user, env) {
  const body = await readJson(request, 2_000_000);
  const taskId = clean(body.taskId);
  const data = await loadData(db);
  const task = data.tasks.find((item) => item.id === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (user.role !== "Coordinador" && task.ownerId !== user.id) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  if (!taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled) }, 409);
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
  const storageProvider = externalStorageProvider(env);
  await db
    .prepare(
      "INSERT INTO evidence_files (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    )
    .bind(fileId, task.id, task.ownerId, user.id, fileName, mimeType, storageProvider ? "" : fileMatch[2], createdAt)
    .run();
  if (storageProvider) {
    try {
      const file = { id: fileId, fileName, mimeType };
      if (storageProvider === R2_STORAGE_PROVIDER) {
        await uploadEvidenceToR2(env, db, task, file, fileMatch[2]);
      } else {
        await uploadEvidenceToOneDrive(env, db, task, file, fileMatch[2]);
      }
    } catch (error) {
      await db.prepare("DELETE FROM evidence_files WHERE id = ?").bind(fileId).run();
      console.error("External evidence upload", error);
      return json({ ok: false, message: "No se pudo guardar el archivo en la nube. Intenta nuevamente." }, 503);
    }
  }

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
  if (!taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled) }, 409);
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
  const row = await db.prepare("SELECT task_id, owner_id, submitted_by_id FROM evidence_files WHERE id = ?").bind(fileId).first();
  if (!row) return json({ ok: false, message: "La carga ya no existe." }, 404);
  if (user.role !== "Coordinador" && row.submitted_by_id !== user.id && row.owner_id !== user.id) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === clean(row.task_id));
  if (!task || !taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)) {
    return json({
      ok: false,
      message: task
        ? evidenceUploadWindowError(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)
        : "La tarea ya no existe."
    }, task ? 409 : 404);
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

async function completeEvidenceUpload(request, db, user, env) {
  const body = await readJson(request);
  const fileId = clean(body.fileId);
  const expectedChunks = Number(body.chunkCount);
  const row = await db.prepare("SELECT * FROM evidence_files WHERE id = ?").bind(fileId).first();
  if (!row) return json({ ok: false, message: "La carga ya no existe." }, 404);
  if (user.role !== "Coordinador" && row.submitted_by_id !== user.id && row.owner_id !== user.id) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === clean(row.task_id));
  if (!task || !taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)) {
    return json({
      ok: false,
      message: task
        ? evidenceUploadWindowError(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)
        : "La tarea ya no existe."
    }, task ? 409 : 404);
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
  const storageProvider = externalStorageProvider(env);
  if (storageProvider) {
    const chunks = await db
      .prepare("SELECT chunk_base64 FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC")
      .bind(fileId)
      .all();
    const fileBase64 = (chunks.results || []).map((item) => item.chunk_base64).join("");
    try {
      const file = { id: row.id, fileName: row.file_name, mimeType: row.mime_type };
      if (storageProvider === R2_STORAGE_PROVIDER) {
        await uploadEvidenceToR2(env, db, task, file, fileBase64);
      } else {
        await uploadEvidenceToOneDrive(env, db, task, file, fileBase64);
      }
      await db.prepare("DELETE FROM evidence_file_chunks WHERE file_id = ?").bind(fileId).run();
    } catch (error) {
      console.error("External chunked evidence upload", error);
      return json({ ok: false, message: "No se pudo guardar el archivo en la nube. La carga puede reintentarse." }, 503);
    }
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

async function evidenceFile(db, user, fileId, env) {
  const row = await db.prepare("SELECT * FROM evidence_files WHERE id = ?").bind(clean(fileId)).first();
  if (!row) return json({ ok: false, message: "Archivo no encontrado." }, 404);
  let observerAllowed = false;
  if (isObserverUser(user)) {
    const data = await loadData(db);
    const task = data.tasks.find((item) => clean(item.id) === clean(row.task_id));
    observerAllowed = Boolean(task);
  }
  const allowed =
    user.role === "Coordinador" || observerAllowed || row.submitted_by_id === user.id || row.owner_id === user.id;
  if (!allowed) return json({ ok: false, message: "No tienes acceso a este archivo." }, 403);

  const storage = await db.prepare("SELECT * FROM evidence_storage WHERE file_id = ?").bind(row.id).first();
  if (storage?.provider === R2_STORAGE_PROVIDER) return r2FileResponse(env, storage, row);
  if (storage?.provider === MICROSOFT_STORAGE_PROVIDER) return oneDriveFileResponse(env, storage, row);

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
  const disposition =
    String(row.mime_type).startsWith("image/") ||
    String(row.mime_type).startsWith("video/") ||
    row.mime_type === "application/pdf"
      ? "inline"
      : "attachment";
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

async function createTask(request, db, user, context) {
  const body = await readJson(request);
  const submitted = body.task || {};
  const taskId = clean(submitted.id) || crypto.randomUUID();
  const title = clean(submitted.title).slice(0, 140);
  const ownerId = user.role === "Coordinador" ? clean(submitted.ownerId) : user.id;
  const dueDate = clean(submitted.dueDate);
  const startTime = clean(submitted.startTime);
  const endTime = clean(submitted.endTime);
  if (!title || title.length < 2) return json({ ok: false, message: "Escribe el nombre de la tarea." }, 400);

  const owner = await db
    .prepare("SELECT id, name, role, status, access_level FROM users WHERE id = ? AND status = 'Activo'")
    .bind(ownerId)
    .first();
  if (!owner || isObserverUser(owner)) return json({ ok: false, message: "El responsable no esta disponible para tareas." }, 400);
  if (user.role !== "Coordinador" && clean(owner.id) !== user.id) {
    return json({ ok: false, message: "Solo puedes crear tareas para tu propio usuario." }, 403);
  }

  const data = await loadData(db);
  const existing = data.tasks.find((task) => clean(task.id) === taskId);
  if (existing) {
    if (clean(existing.createdById) !== user.id || clean(existing.ownerId) !== ownerId) {
      return json({ ok: false, message: "El identificador de la tarea ya esta en uso." }, 409);
    }
    return stateResponse(db, user, data);
  }
  if (
    !validWorkSchedule(
      dueDate,
      startTime,
      endTime,
      breakSettingsForUser(data, ownerId, dueDate),
      owner.name,
      workScheduleEndForUser(data, ownerId, dueDate, owner.name)
    )
  ) {
    const scheduleEnd = workScheduleEndForUser(data, ownerId, dueDate, owner.name);
    const schedule = breakSettingsForUser(data, ownerId, dueDate);
    return json({
      ok: false,
      message: `El horario no esta disponible. Jornada hasta ${scheduleEnd}${new Date(`${dueDate}T12:00:00Z`).getUTCDay() === 6 ? " sin break" : ` y break ${schedule.breakStart}-${schedule.breakEnd}`}.`
    }, 409);
  }
  if (hasTaskConflict(data.tasks, ownerId, dueDate, startTime, endTime)) {
    return json({ ok: false, message: "Ese horario ya esta ocupado por otra tarea." }, 409);
  }

  const createdAt = Date.now();
  const task = {
    id: taskId,
    title,
    ownerId,
    createdById: user.id,
    category: normalizeTaskCategory(submitted.category),
    priority: ["Alta", "Media", "Baja"].includes(submitted.priority) ? submitted.priority : "Media",
    dueDate,
    startTime,
    endTime,
    product: clean(submitted.product).slice(0, 120) || "GENERAL",
    description: clean(submitted.description).slice(0, 1500),
    status: "Pendiente",
    createdAt,
    history: [{ type: "Asignacion", toId: ownerId, byId: user.id, reason: "Tarea creada", at: createdAt }],
    evidence: [],
    reminders: []
  };
  data.tasks.unshift(task);
  await saveData(db, data);
  if (ownerId !== user.id) {
    const notifiedUsers = await queueNotifications(db, [{
      userId: ownerId,
      title: "Nueva tarea asignada",
      body: `${title} | ${dueDate} ${startTime}-${endTime}`,
      url: `/?view=tasksView&task=${encodeURIComponent(taskId)}`,
      sourceKey: `task:create:${taskId}:${ownerId}`
    }]);
    if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  }
  return stateResponse(db, user, data);
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
  if (!taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, dateIsoInLima(Date.now()), data.lateEvidenceUploadsEnabled) }, 409);
  }

  const result = "Realizado";
  const notes = clean(submittedEvidence.notes).slice(0, 3000);
  const category = clean(task.category).toUpperCase();
  const store =
    clean(submittedEvidence.store).slice(0, 180) ||
    (["PDP", "PDV"].includes(category) ? "No especificada" : "No aplica");
  const product = clean(submittedEvidence.product).slice(0, 180);
  const trainingTopic = clean(submittedEvidence.trainingTopic).slice(0, 300);
  const links = [];
  for (const submittedLink of Array.isArray(submittedEvidence.links) ? submittedEvidence.links.slice(0, 10) : []) {
    const link = clean(submittedLink).slice(0, 1000);
    let parsed;
    try {
      parsed = new URL(link);
    } catch {
      return json({ ok: false, message: "Uno de los enlaces del sustento no es valido." }, 400);
    }
    if (!["http:", "https:"].includes(parsed.protocol)) {
      return json({ ok: false, message: "Los enlaces deben utilizar http:// o https://." }, 400);
    }
    links.push(parsed.href);
  }
  if (normalizeTaskCategory(task.category) === "Entrenamiento" && !trainingTopic) {
    return json({ ok: false, message: "Completa el tema de la capacitacion." }, 400);
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
  const submittedDate = dateIsoInLima(submittedAt);
  const isLateSubmission = clean(task.dueDate) < submittedDate;
  const individualLateAuthorization = isLateSubmission && lateEvidenceAuthorizationActive(task)
    ? { ...task.lateEvidenceAuthorization, scope: "individual" }
    : null;
  const latestGlobalPolicy = Array.isArray(data.lateEvidencePolicyHistory)
    ? data.lateEvidencePolicyHistory.find((entry) => entry?.enabled)
    : null;
  const lateAuthorization = individualLateAuthorization || (isLateSubmission && data.lateEvidenceUploadsEnabled
    ? {
        enabled: true,
        scope: "global",
        reason: "Carga de sustentos anteriores habilitada por Pablo",
        authorizedById: clean(latestGlobalPolicy?.byId),
        authorizedAt: Number(latestGlobalPolicy?.at || submittedAt)
      }
    : null);
  const evidence = {
    id: evidenceId,
    submittedById: user.id,
    submittedAt,
    store,
    product,
    trainingTopic: normalizeTaskCategory(task.category) === "Entrenamiento" ? trainingTopic : "",
    result,
    notes,
    links,
    files,
    cloudPath: clean(submittedEvidence.cloudPath).slice(0, 800),
    review: "Pendiente"
  };
  if (lateAuthorization) evidence.lateAuthorization = lateAuthorization;
  task.evidence.push(evidence);
  task.history = Array.isArray(task.history) ? task.history : [];
  task.status = "En revision";
  task.blockedReason = "";
  task.blockedAt = 0;
  if (individualLateAuthorization) {
    task.lateEvidenceAuthorization = {
      ...task.lateEvidenceAuthorization,
      enabled: false,
      usedAt: submittedAt,
      usedEvidenceId: evidenceId
    };
  }
  if (lateAuthorization) {
    task.history.push({
      type: "SustentoTardio",
      byId: user.id,
      reason: lateAuthorization.reason,
      at: submittedAt
    });
  }
  task.history.push({
    type: "Sustento",
    byId: user.id,
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

async function authorizeLateTaskEvidence(request, db, user, context) {
  if (user.role !== "Coordinador") {
    return json({ ok: false, message: "Solo el coordinador puede autorizar una carga tardia." }, 403);
  }
  const body = await readJson(request);
  const taskId = clean(body.taskId);
  const reason = clean(body.reason).slice(0, 1000);
  if (!taskId || !reason) return json({ ok: false, message: "Indica la tarea y el motivo de la autorizacion." }, 400);
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  const today = dateIsoInLima(Date.now());
  if (clean(task.dueDate) >= today) {
    return json({ ok: false, message: "La autorizacion tardia solo se usa para tareas de dias anteriores." }, 409);
  }
  if (task.status === "Cumplida") {
    return json({ ok: false, message: "La tarea ya fue aprobada y no necesita otro sustento." }, 409);
  }
  const authorizedAt = Date.now();
  task.lateEvidenceAuthorization = {
    enabled: true,
    validDate: today,
    reason,
    authorizedById: user.id,
    authorizedAt
  };
  task.history = Array.isArray(task.history) ? task.history : [];
  task.history.push({
    type: "AutorizacionSustentoTardio",
    byId: user.id,
    reason,
    validDate: today,
    at: authorizedAt
  });
  await saveData(db, data);
  const notifiedUsers = await queueNotifications(db, [{
    userId: task.ownerId,
    title: "Carga tardia autorizada",
    body: `${clean(task.title)} | Disponible por un envio durante el dia de hoy.`,
    url: `/?view=tasksView&task=${encodeURIComponent(taskId)}`,
    sourceKey: `late-evidence:${taskId}:${authorizedAt}`
  }]);
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
  const reviewNote = clean(body.reviewNote).slice(0, 1500);
  if (!["Cumplida", "Observada"].includes(status)) {
    return json({ ok: false, message: "La revision indicada no es valida." }, 400);
  }
  if (status === "Observada" && !reviewNote) {
    return json({ ok: false, message: "Indica el motivo del rechazo." }, 400);
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
  latestEvidence.reviewNote = reviewNote || "Sustento aprobado";
  task.completedAt = status === "Cumplida" ? reviewedAt : 0;
  task.history = Array.isArray(task.history) ? task.history : [];
  task.history.push({
    type: status === "Cumplida" ? "Aprobacion" : "Rechazo",
    byId: user.id,
    reason: reviewNote || (status === "Cumplida" ? "Sustento aprobado y tarea completada" : "Sustento rechazado por el coordinador"),
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

async function deleteOneDriveEvidenceItems(env, items) {
  if (!microsoftStorageConfigured(env)) return;
  for (const item of items) {
    try {
      const response = await graphFetch(
        env,
        `/drives/${encodeURIComponent(item.drive_id)}/items/${encodeURIComponent(item.drive_item_id)}`,
        { method: "DELETE" }
      );
      if (!response.ok && response.status !== 404) {
        console.error("OneDrive evidence deletion", response.status);
      }
    } catch (error) {
      console.error("OneDrive evidence deletion", error);
    }
  }
}

async function deleteR2EvidenceItems(env, items) {
  if (!r2StorageEnabled(env) || !items.length) return;
  const keys = items.map((item) => clean(item.drive_item_id)).filter(Boolean);
  if (!keys.length) return;
  try {
    await env.EVIDENCE_BUCKET.delete(keys);
  } catch (error) {
    console.error("R2 evidence deletion", error);
  }
}

async function deleteExternalEvidenceItems(env, items) {
  const r2Items = items.filter((item) => item.provider === R2_STORAGE_PROVIDER);
  const oneDriveItems = items.filter((item) => item.provider === MICROSOFT_STORAGE_PROVIDER);
  await Promise.all([deleteR2EvidenceItems(env, r2Items), deleteOneDriveEvidenceItems(env, oneDriveItems)]);
}

async function deleteTaskEvidenceFiles(db, taskId, env, context) {
  const storedItems = await db
    .prepare(
      `SELECT storage.provider, storage.drive_id, storage.drive_item_id
       FROM evidence_storage storage
       INNER JOIN evidence_files files ON files.id = storage.file_id
       WHERE files.task_id = ?`
    )
    .bind(taskId)
    .all();
  await db
    .prepare("DELETE FROM evidence_file_chunks WHERE file_id IN (SELECT id FROM evidence_files WHERE task_id = ?)")
    .bind(taskId)
    .run();
  await db
    .prepare("DELETE FROM evidence_storage WHERE file_id IN (SELECT id FROM evidence_files WHERE task_id = ?)")
    .bind(taskId)
    .run();
  await db.prepare("DELETE FROM evidence_files WHERE task_id = ?").bind(taskId).run();
  if (storedItems.results?.length) {
    context.waitUntil(deleteExternalEvidenceItems(env, storedItems.results));
  }
}

async function deleteTaskAndArchive(request, db, user, context, env) {
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
  await deleteTaskEvidenceFiles(db, taskId, env, context);
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

async function saveOvertimeSchedule(request, db, user, context) {
  if (user.role !== "Coordinador") {
    return json({ ok: false, message: "Los trainers deben solicitar autorizacion a Pablo." }, 403);
  }
  const body = await readJson(request, 20_000);
  const targetUserId = clean(body.userId);
  const dateValue = clean(body.date);
  const endTime = clean(body.endTime);
  const reason = clean(body.reason).slice(0, 300);
  const target = await db
    .prepare("SELECT id, name, email, zone, role, status, created_at, access_level FROM users WHERE id = ? AND status = 'Activo'")
    .bind(targetUserId)
    .first();
  if (!target || target.role !== "Trainer" || isObserverUser(target)) return json({ ok: false, message: "Selecciona un trainer activo." }, 400);
  if (user.role !== "Coordinador" && target.id !== user.id) {
    return json({ ok: false, message: "Solo puedes ampliar tu propia jornada." }, 403);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateValue) || new Date(`${dateValue}T12:00:00Z`).getUTCDay() === 0) {
    return json({ ok: false, message: "Selecciona un dia laborable de lunes a sabado." }, 400);
  }
  const baseEnd = baseWorkdayEnd(dateValue, target.name);
  if (
    !/^\d{2}:\d{2}$/.test(endTime) ||
    timeToMinutes(endTime) < timeToMinutes(baseEnd) ||
    timeToMinutes(endTime) > timeToMinutes("23:45")
  ) {
    return json({ ok: false, message: `La salida debe estar entre ${baseEnd} y 23:45.` }, 400);
  }
  if (!reason) return json({ ok: false, message: "Escribe el motivo del cambio de jornada." }, 400);

  const data = await loadData(db);
  const updatedAt = Date.now();
  data.workScheduleByUserDate = {
    ...(data.workScheduleByUserDate || {}),
    [target.id]: {
      ...(data.workScheduleByUserDate?.[target.id] || {}),
      [dateValue]: {
        endTime,
        reason,
        updatedById: user.id,
        updatedByName: user.name,
        updatedAt
      }
    }
  };
  data.overtimeRequests = (data.overtimeRequests || []).map((item) =>
    item.status === "Pendiente" && item.userId === target.id && item.date === dateValue
      ? {
          ...item,
          status: item.endTime === endTime ? "Aprobada" : "Rechazada",
          reviewedAt: updatedAt,
          reviewedById: user.id,
          reviewedByName: user.name,
          reviewNote:
            item.endTime === endTime
              ? "Aprobada al configurar la jornada directamente."
              : "Resuelta al configurar un horario diferente."
        }
      : item
  );
  await saveData(db, data);

  const restored = endTime === baseEnd;
  const notifications = [{
    userId: target.id,
    title: restored ? "Jornada normal restablecida" : "Jornada ampliada",
    body: `${target.name} | ${dateValue} | salida ${endTime} | ${reason}`,
    url: `/?view=tasksView`,
    sourceKey: `overtime:${target.id}:${dateValue}:${endTime}:${updatedAt}`
  }];
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function requestOvertimeSchedule(request, db, user, context) {
  if (user.role !== "Trainer") return json({ ok: false, message: "Solo los trainers pueden enviar esta solicitud." }, 403);
  const body = await readJson(request, 20_000);
  const dateValue = clean(body.date);
  const endTime = clean(body.endTime);
  const reason = clean(body.reason).slice(0, 300);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateValue) || new Date(`${dateValue}T12:00:00Z`).getUTCDay() === 0) {
    return json({ ok: false, message: "Selecciona un dia laborable de lunes a sabado." }, 400);
  }
  const baseEnd = baseWorkdayEnd(dateValue, user.name);
  if (
    !/^\d{2}:\d{2}$/.test(endTime) ||
    timeToMinutes(endTime) < timeToMinutes(baseEnd) ||
    timeToMinutes(endTime) > timeToMinutes("23:45")
  ) {
    return json({ ok: false, message: `La salida debe estar entre ${baseEnd} y 23:45.` }, 400);
  }
  if (!reason) return json({ ok: false, message: "Escribe el motivo de la solicitud." }, 400);

  const data = await loadData(db);
  const currentEnd = workScheduleEndForUser(data, user.id, dateValue, user.name);
  if (endTime === currentEnd) return json({ ok: false, message: `Tu jornada ya termina a las ${endTime}.` }, 400);
  const requestedAt = Date.now();
  const existingIndex = (data.overtimeRequests || []).findIndex(
    (item) => item.userId === user.id && item.date === dateValue && item.status === "Pendiente"
  );
  const previous = existingIndex >= 0 ? data.overtimeRequests[existingIndex] : null;
  const overtimeRequest = {
    id: previous?.id || crypto.randomUUID(),
    userId: user.id,
    userName: user.name,
    date: dateValue,
    endTime,
    reason,
    status: "Pendiente",
    requestedAt,
    reviewedAt: 0,
    reviewedById: "",
    reviewedByName: "",
    reviewNote: ""
  };
  if (existingIndex >= 0) data.overtimeRequests.splice(existingIndex, 1, overtimeRequest);
  else data.overtimeRequests.unshift(overtimeRequest);
  await saveData(db, data);

  const coordinators = await db.prepare("SELECT id FROM users WHERE role = 'Coordinador' AND status = 'Activo'").all();
  const notifications = (coordinators.results || []).map((coordinator) => ({
    userId: coordinator.id,
    title: "Solicitud de horas extra",
    body: `${user.name} solicita salida ${endTime} para ${dateValue}. ${reason}`,
    url: "/?view=tasksView",
    sourceKey: `overtime-request:${overtimeRequest.id}:${requestedAt}`
  }));
  const notifiedUsers = await queueNotifications(db, notifications);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function reviewOvertimeSchedule(request, db, user, context) {
  if (user.role !== "Coordinador") return json({ ok: false, message: "Solo el coordinador puede autorizar horas extra." }, 403);
  const body = await readJson(request, 20_000);
  const requestId = clean(body.requestId);
  const decision = clean(body.decision);
  const note = clean(body.note).slice(0, 300);
  if (!["Aprobar", "Rechazar"].includes(decision)) return json({ ok: false, message: "Selecciona una decision valida." }, 400);
  if (decision === "Rechazar" && !note) return json({ ok: false, message: "Escribe el motivo del rechazo." }, 400);

  const data = await loadData(db);
  const requestIndex = (data.overtimeRequests || []).findIndex((item) => item.id === requestId && item.status === "Pendiente");
  if (requestIndex < 0) return json({ ok: false, message: "La solicitud ya fue revisada o no existe." }, 404);
  const overtimeRequest = data.overtimeRequests[requestIndex];
  const target = await db
    .prepare("SELECT id, name, email, zone, role, status, created_at, access_level FROM users WHERE id = ? AND status = 'Activo'")
    .bind(overtimeRequest.userId)
    .first();
  if (!target || target.role !== "Trainer" || isObserverUser(target)) return json({ ok: false, message: "El trainer ya no esta activo." }, 400);
  const baseEnd = baseWorkdayEnd(overtimeRequest.date, target.name);
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(overtimeRequest.date) ||
    new Date(`${overtimeRequest.date}T12:00:00Z`).getUTCDay() === 0 ||
    !/^\d{2}:\d{2}$/.test(overtimeRequest.endTime) ||
    timeToMinutes(overtimeRequest.endTime) < timeToMinutes(baseEnd) ||
    timeToMinutes(overtimeRequest.endTime) > timeToMinutes("23:45")
  ) {
    return json({ ok: false, message: "El horario solicitado ya no es valido." }, 400);
  }

  const reviewedAt = Date.now();
  if (decision === "Aprobar") {
    data.workScheduleByUserDate = {
      ...(data.workScheduleByUserDate || {}),
      [target.id]: {
        ...(data.workScheduleByUserDate?.[target.id] || {}),
        [overtimeRequest.date]: {
          endTime: overtimeRequest.endTime,
          reason: overtimeRequest.reason,
          updatedById: user.id,
          updatedByName: user.name,
          updatedAt: reviewedAt
        }
      }
    };
  }
  data.overtimeRequests[requestIndex] = {
    ...overtimeRequest,
    status: decision === "Aprobar" ? "Aprobada" : "Rechazada",
    reviewedAt,
    reviewedById: user.id,
    reviewedByName: user.name,
    reviewNote: note
  };
  await saveData(db, data);

  const notification = {
    userId: target.id,
    title: decision === "Aprobar" ? "Horas extra aprobadas" : "Solicitud de horas extra rechazada",
    body:
      decision === "Aprobar"
        ? `${overtimeRequest.date} | jornada autorizada hasta ${overtimeRequest.endTime}${note ? ` | ${note}` : ""}`
        : `${overtimeRequest.date} | ${note}`,
    url: "/?view=tasksView",
    sourceKey: `overtime-review:${overtimeRequest.id}:${decision}:${reviewedAt}`
  };
  const notifiedUsers = await queueNotifications(db, [notification]);
  if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  return stateResponse(db, user, data);
}

async function stateResponse(db, user, loadedData = null) {
  const data = loadedData || (await loadData(db));
  const coordinator = user.role === "Coordinador";
  const observer = isObserverUser(user);
  const canViewAll = coordinator || observer;
  const userRows = canViewAll
    ? await db.prepare("SELECT id, name, email, zone, role, status, created_at, access_level FROM users ORDER BY created_at ASC").all()
    : await db
        .prepare("SELECT id, name, email, zone, role, status, created_at, access_level FROM users WHERE status = 'Activo' ORDER BY role ASC, name ASC")
        .all();
  const users = (userRows.results || []).map(publicUser);
  const state = {
    ...EMPTY_DATA,
    ...data,
    activeUserId: user.id,
    users,
    tasks: canViewAll ? data.tasks : data.tasks.filter((task) => task.ownerId === user.id),
    deletedTasks: canViewAll
      ? data.deletedTasks
      : (data.deletedTasks || []).filter((task) => task.ownerId === user.id),
    workScheduleByUserDate: canViewAll
      ? data.workScheduleByUserDate
      : { [user.id]: data.workScheduleByUserDate?.[user.id] || {} },
    overtimeRequests: canViewAll
      ? data.overtimeRequests
      : (data.overtimeRequests || []).filter((request) => request.userId === user.id),
    announcements: canViewAll
      ? data.announcements
      : data.announcements.filter((item) => item.audience === "all" || item.targetId === user.id),
    supportRequests: canViewAll
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
  const activeUserRows = await db.prepare("SELECT id, name, role, access_level FROM users WHERE status = 'Activo'").all();
  const activeUsers = activeUserRows.results || [];
  const activeUserIds = new Set(activeUsers.map((row) => clean(row.id)));
  const userNamesById = new Map(activeUsers.map((row) => [row.id, row.name]));
  const personalBreak = submitted.breakSettingsByUser?.[user.id];
  if (validWorkSettings(personalBreak)) {
    current.breakSettingsByUser = {
      ...(current.breakSettingsByUser || {}),
      [user.id]: normalizeWorkSettings(personalBreak)
    };
  }
  const personalBreakDates = submitted.breakSettingsByUserDate?.[user.id];
  if (personalBreakDates && typeof personalBreakDates === "object" && !Array.isArray(personalBreakDates)) {
    const normalizedDates = normalizeBreakSettingsByUserDate({ [user.id]: personalBreakDates });
    const conflictFreeDates = Object.fromEntries(
      Object.entries(normalizedDates[user.id] || {}).filter(
        ([dateValue, settings]) =>
          !hasTaskConflict(current.tasks, user.id, dateValue, settings.breakStart, settings.breakEnd)
      )
    );
    current.breakSettingsByUserDate = {
      ...(current.breakSettingsByUserDate || {}),
      [user.id]: {
        ...(current.breakSettingsByUserDate?.[user.id] || {}),
        ...conflictFreeDates
      }
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
          if (!previous && (!activeUserIds.has(clean(normalizedTask.ownerId)) || clean(normalizedTask.title).length < 2)) return null;
          const previousHistory = Array.isArray(previous?.history) ? previous.history : [];
          const submittedHistory = Array.isArray(normalizedTask.history) ? normalizedTask.history : previousHistory;
          const appendOnlyHistory = historyIsAppendOnly(previousHistory, submittedHistory);
          const lastHistory = submittedHistory.at(-1);
          if (previous) {
            const ownerChanged = clean(previous.ownerId) !== clean(normalizedTask.ownerId);
            const scheduleChanged =
              previous.dueDate !== normalizedTask.dueDate ||
              previous.startTime !== normalizedTask.startTime ||
              previous.endTime !== normalizedTask.endTime;
            const titleChanged = clean(previous.title) !== clean(normalizedTask.title);
            const validReassignmentAudit =
              !ownerChanged ||
              (appendOnlyHistory &&
                lastHistory?.type === "Reasignacion" &&
                clean(lastHistory.fromId) === clean(previous.ownerId) &&
                clean(lastHistory.toId) === clean(normalizedTask.ownerId) &&
                clean(lastHistory.byId) === user.id &&
                Boolean(clean(lastHistory.reason)));
            const validScheduleAudit =
              !scheduleChanged ||
              (appendOnlyHistory &&
                lastHistory?.type === "Reprogramacion" &&
                clean(lastHistory.byId) === user.id &&
                clean(lastHistory.fromDate) === clean(previous.dueDate) &&
                clean(lastHistory.toDate) === clean(normalizedTask.dueDate) &&
                clean(lastHistory.fromStartTime) === clean(previous.startTime || WORKDAY_START) &&
                clean(lastHistory.fromEndTime) === clean(previous.endTime || "09:30") &&
                clean(lastHistory.toStartTime) === clean(normalizedTask.startTime) &&
                clean(lastHistory.toEndTime) === clean(normalizedTask.endTime) &&
                Boolean(clean(lastHistory.reason)));
            const validTitleAudit =
              !titleChanged ||
              (appendOnlyHistory &&
                !ownerChanged &&
                !scheduleChanged &&
                clean(normalizedTask.title).length >= 2 &&
                lastHistory?.type === "CorreccionNombre" &&
                clean(lastHistory.byId) === user.id &&
                clean(lastHistory.fromTitle) === clean(previous.title) &&
                clean(lastHistory.toTitle) === clean(normalizedTask.title) &&
                Boolean(clean(lastHistory.reason)));
            if (!appendOnlyHistory || !validReassignmentAudit || !validScheduleAudit || !validTitleAudit) return previous;
            normalizedTask.title = titleChanged ? clean(normalizedTask.title).slice(0, 140) : previous.title;
            normalizedTask.history = submittedHistory;
          }
          const ownerOrScheduleChanged =
            !previous ||
            clean(previous.ownerId) !== clean(normalizedTask.ownerId) ||
            previous.dueDate !== normalizedTask.dueDate ||
            previous.startTime !== normalizedTask.startTime ||
            previous.endTime !== normalizedTask.endTime;
          if (
            ownerOrScheduleChanged &&
            (
              !validWorkSchedule(
                clean(normalizedTask.dueDate),
                clean(normalizedTask.startTime),
                clean(normalizedTask.endTime),
                breakSettingsForUser(current, normalizedTask.ownerId, clean(normalizedTask.dueDate)),
                userNamesById.get(clean(normalizedTask.ownerId)),
                workScheduleEndForUser(
                  current,
                  normalizedTask.ownerId,
                  clean(normalizedTask.dueDate),
                  userNamesById.get(clean(normalizedTask.ownerId))
                )
              ) ||
              hasTaskConflict(
                current.tasks,
                clean(normalizedTask.ownerId),
                clean(normalizedTask.dueDate),
                clean(normalizedTask.startTime),
                clean(normalizedTask.endTime),
                clean(normalizedTask.id)
              )
            )
          ) {
            return previous || null;
          }
          if (previous) {
            normalizedTask.evidence = previous.evidence;
            normalizedTask.lateEvidenceAuthorization = previous.lateEvidenceAuthorization || null;
            normalizedTask.completedAt = previous.completedAt || 0;
          } else {
            normalizedTask.evidence = [];
            normalizedTask.lateEvidenceAuthorization = null;
            normalizedTask.completedAt = 0;
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
    const activeTrainerIds = new Set(
      activeUsers.filter((row) => row.role === "Trainer" && !isObserverUser(row)).map((row) => row.id)
    );
    current.tasks = current.tasks.map((task) => {
      if (task.ownerId !== user.id) return task;
      const next = submittedTasks.get(task.id);
      if (!next) return task;
      const requestedOwnerId = clean(next.ownerId);
      const ownerChanged = requestedOwnerId && requestedOwnerId !== task.ownerId && activeTrainerIds.has(requestedOwnerId);
      const nextHistory = Array.isArray(next.history) ? next.history : task.history;
      const appendOnlyHistory = historyIsAppendOnly(task.history, nextHistory);
      const lastHistory = nextHistory?.at(-1);
      const validReassignment =
        ownerChanged &&
        appendOnlyHistory &&
        lastHistory?.type === "Reasignacion" &&
        lastHistory.fromId === user.id &&
        lastHistory.toId === requestedOwnerId &&
        lastHistory.byId === user.id &&
        clean(lastHistory.reason) &&
        validWorkSchedule(
          clean(task.dueDate),
          clean(task.startTime),
          clean(task.endTime),
          breakSettingsForUser(current, requestedOwnerId, clean(task.dueDate)),
          userNamesById.get(requestedOwnerId),
          workScheduleEndForUser(current, requestedOwnerId, clean(task.dueDate), userNamesById.get(requestedOwnerId))
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
        appendOnlyHistory &&
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
        validWorkSchedule(
          requestedDate,
          requestedStart,
          requestedEnd,
          breakSettingsForUser(current, user.id, requestedDate),
          user.name,
          workScheduleEndForUser(current, user.id, requestedDate, user.name)
        ) &&
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
      const requestedTitle = clean(next.title).slice(0, 140);
      const titleChanged = requestedTitle && requestedTitle !== clean(task.title);
      const validRename =
        !ownerChanged &&
        !scheduleChanged &&
        titleChanged &&
        appendOnlyHistory &&
        lastHistory?.type === "CorreccionNombre" &&
        clean(lastHistory.byId) === user.id &&
        clean(lastHistory.fromTitle) === clean(task.title) &&
        clean(lastHistory.toTitle) === requestedTitle &&
        Boolean(clean(lastHistory.reason));
      const validStart =
        task.status === "Pendiente" &&
        requestedStatus === "En proceso" &&
        appendOnlyHistory &&
        lastHistory?.type === "Inicio" &&
        clean(lastHistory.byId) === user.id;
      const remindersChanged = JSON.stringify(next.reminders || []) !== JSON.stringify(task.reminders || []);
      const validReminderUpdate =
        remindersChanged &&
        appendOnlyHistory &&
        lastHistory?.type === "Recordatorio" &&
        clean(lastHistory.byId) === user.id &&
        validReminderList(next.reminders, user.id, task.ownerId);
      return {
        ...task,
        title: validRename ? requestedTitle : task.title,
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
          validReassignment || validReschedule || validRename || validStart || validReminderUpdate
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
        !validWorkSchedule(
          dueDate,
          startTime,
          endTime,
          breakSettingsForUser(current, user.id, dueDate),
          user.name,
          workScheduleEndForUser(current, user.id, dueDate, user.name)
        ) ||
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
    mp4: "video/mp4",
    mov: "video/quicktime",
    webm: "video/webm",
    m4v: "video/x-m4v",
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

function normalizeBreakSettingsByUserDate(settingsByUserDate) {
  if (!settingsByUserDate || typeof settingsByUserDate !== "object" || Array.isArray(settingsByUserDate)) return {};
  return Object.fromEntries(
    Object.entries(settingsByUserDate)
      .filter(([userId, dates]) => clean(userId) && dates && typeof dates === "object" && !Array.isArray(dates))
      .map(([userId, dates]) => [
        clean(userId),
        Object.fromEntries(
          Object.entries(dates)
            .filter(([dateValue, settings]) => /^\d{4}-\d{2}-\d{2}$/.test(dateValue) && validWorkSettings(settings))
            .map(([dateValue, settings]) => [dateValue, normalizeWorkSettings(settings)])
        )
      ])
  );
}

function normalizeWorkScheduleByUserDate(settingsByUserDate) {
  if (!settingsByUserDate || typeof settingsByUserDate !== "object" || Array.isArray(settingsByUserDate)) return {};
  return Object.fromEntries(
    Object.entries(settingsByUserDate)
      .filter(([userId, dates]) => clean(userId) && dates && typeof dates === "object" && !Array.isArray(dates))
      .map(([userId, dates]) => [
        clean(userId),
        Object.fromEntries(
          Object.entries(dates)
            .filter(
              ([dateValue, settings]) =>
                /^\d{4}-\d{2}-\d{2}$/.test(dateValue) &&
                settings &&
                /^\d{2}:\d{2}$/.test(clean(settings.endTime)) &&
                timeToMinutes(settings.endTime) <= timeToMinutes("23:45")
            )
            .map(([dateValue, settings]) => [
              dateValue,
              {
                endTime: clean(settings.endTime),
                reason: clean(settings.reason).slice(0, 300),
                updatedById: clean(settings.updatedById),
                updatedByName: clean(settings.updatedByName),
                updatedAt: Number(settings.updatedAt || 0)
              }
            ])
        )
      ])
  );
}

function normalizeOvertimeRequests(requests) {
  if (!Array.isArray(requests)) return [];
  return requests
    .filter(
      (request) =>
        request &&
        clean(request.id) &&
        clean(request.userId) &&
        /^\d{4}-\d{2}-\d{2}$/.test(clean(request.date)) &&
        /^\d{2}:\d{2}$/.test(clean(request.endTime))
    )
    .map((request) => ({
      id: clean(request.id),
      userId: clean(request.userId),
      userName: clean(request.userName),
      date: clean(request.date),
      endTime: clean(request.endTime),
      reason: clean(request.reason).slice(0, 300),
      status: ["Pendiente", "Aprobada", "Rechazada"].includes(request.status) ? request.status : "Pendiente",
      requestedAt: Number(request.requestedAt || 0),
      reviewedAt: Number(request.reviewedAt || 0),
      reviewedById: clean(request.reviewedById),
      reviewedByName: clean(request.reviewedByName),
      reviewNote: clean(request.reviewNote).slice(0, 300)
    }));
}

function breakSettingsForUser(data, userId, dateValue = "") {
  const daily = dateValue ? data.breakSettingsByUserDate?.[clean(userId)]?.[dateValue] : null;
  if (validWorkSettings(daily)) return normalizeWorkSettings(daily);
  const personal = data.breakSettingsByUser?.[clean(userId)];
  return validWorkSettings(personal) ? normalizeWorkSettings(personal) : normalizeWorkSettings(data.workSettings);
}

function normalizeTaskCategory(category) {
  if (category === "PDV") return "PDP";
  if (category === "Producto") return "ULG";
  return ["PDP", "Entrenamiento", "ULG", "Reporte", "Coordinacion"].includes(category) ? category : "PDP";
}

function normalizedUserName(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function normalizeLgUpdates(items) {
  if (!Array.isArray(items)) return [];
  return items
    .filter((item) => clean(item?.id) && clean(item?.title) && clean(item?.description))
    .map((item) => ({
      id: clean(item.id),
      line: ["HS", "TV", "AV"].includes(item.line) ? item.line : "HS",
      type: ["Producto", "Lanzamiento", "Tecnologia", "Argumento de venta"].includes(item.type) ? item.type : "Producto",
      title: clean(item.title).slice(0, 140),
      product: clean(item.product).slice(0, 120),
      description: clean(item.description).slice(0, 2000),
      createdById: clean(item.createdById),
      createdByName: clean(item.createdByName),
      createdAt: Number(item.createdAt || 0),
      updatedById: clean(item.updatedById),
      updatedByName: clean(item.updatedByName),
      updatedAt: Number(item.updatedAt || item.createdAt || 0)
    }))
    .slice(0, 500);
}

function baseWorkdayEnd(dateValue, userName = "") {
  const day = new Date(`${dateValue}T12:00:00Z`).getUTCDay();
  const dannySaturday = day === 6 && normalizedUserName(userName) === "DANNY DIOS";
  return day === 6 ? (dannySaturday ? DANNY_SATURDAY_END : SATURDAY_END) : WEEKDAY_END;
}

function workScheduleEndForUser(data, userId, dateValue, userName = "") {
  const baseEnd = baseWorkdayEnd(dateValue, userName);
  const requestedEnd = clean(data.workScheduleByUserDate?.[clean(userId)]?.[dateValue]?.endTime);
  return /^\d{2}:\d{2}$/.test(requestedEnd) &&
    timeToMinutes(requestedEnd) >= timeToMinutes(baseEnd) &&
    timeToMinutes(requestedEnd) <= timeToMinutes("23:45")
    ? requestedEnd
    : baseEnd;
}

function validWorkSchedule(dateValue, startTime, endTime, workSettings = null, userName = "", workEndOverride = "") {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateValue || ""))) return false;
  const day = new Date(`${dateValue}T12:00:00Z`).getUTCDay();
  if (day === 0) return false;
  const start = timeToMinutes(startTime);
  const end = timeToMinutes(endTime);
  const dannySaturday = day === 6 && normalizedUserName(userName) === "DANNY DIOS";
  const workStart = timeToMinutes(dannySaturday ? DANNY_SATURDAY_START : WORKDAY_START);
  const baseEnd = baseWorkdayEnd(dateValue, userName);
  const validOverride =
    /^\d{2}:\d{2}$/.test(clean(workEndOverride)) &&
    timeToMinutes(workEndOverride) >= timeToMinutes(baseEnd) &&
    timeToMinutes(workEndOverride) <= timeToMinutes("23:45");
  const workEnd = timeToMinutes(validOverride ? workEndOverride : baseEnd);
  if (start < workStart || end > workEnd || end <= start) return false;
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
      version: 23,
      workSettings: normalizeWorkSettings(parsed.workSettings),
      breakSettingsByUser: normalizeBreakSettingsByUser(parsed.breakSettingsByUser),
      breakSettingsByUserDate: normalizeBreakSettingsByUserDate(parsed.breakSettingsByUserDate),
      workScheduleByUserDate: normalizeWorkScheduleByUserDate(parsed.workScheduleByUserDate),
      overtimeRequests: normalizeOvertimeRequests(parsed.overtimeRequests),
      tasks: (parsed.tasks || []).map((task) => ({ ...task, category: normalizeTaskCategory(task.category) })),
      deletedTasks: (parsed.deletedTasks || []).map((task) => ({ ...task, category: normalizeTaskCategory(task.category) })),
      lgUpdates: normalizeLgUpdates(parsed.lgUpdates),
      lateEvidenceUploadsEnabled: Boolean(parsed.lateEvidenceUploadsEnabled),
      lateEvidencePolicyHistory: Array.isArray(parsed.lateEvidencePolicyHistory)
        ? parsed.lateEvidencePolicyHistory.slice(0, 100)
        : []
    };
  } catch {
    return structuredClone(EMPTY_DATA);
  }
}

async function saveData(db, data) {
  const payload = {
    version: 23,
    workSettings: normalizeWorkSettings(data.workSettings),
    breakSettingsByUser: normalizeBreakSettingsByUser(data.breakSettingsByUser),
    breakSettingsByUserDate: normalizeBreakSettingsByUserDate(data.breakSettingsByUserDate),
    workScheduleByUserDate: normalizeWorkScheduleByUserDate(data.workScheduleByUserDate),
    overtimeRequests: normalizeOvertimeRequests(data.overtimeRequests),
    registrationRequests: data.registrationRequests || [],
    passwordRecoveryRequests: data.passwordRecoveryRequests || [],
    tasks: data.tasks || [],
    deletedTasks: data.deletedTasks || [],
    announcements: data.announcements || [],
    supportRequests: data.supportRequests || [],
    dailyMotivations: data.dailyMotivations || [],
    lgUpdates: normalizeLgUpdates(data.lgUpdates),
    lateEvidenceUploadsEnabled: Boolean(data.lateEvidenceUploadsEnabled),
    lateEvidencePolicyHistory: Array.isArray(data.lateEvidencePolicyHistory)
      ? data.lateEvidencePolicyHistory.slice(0, 100)
      : []
  };
  await db.prepare("UPDATE app_data SET data = ?, updated_at = ? WHERE id = 1").bind(JSON.stringify(payload), Date.now()).run();
}

function isObserverUser(user) {
  return clean(user?.access_level).toLowerCase() === OBSERVER_ACCESS_LEVEL;
}

function publicUser(row) {
  const observer = isObserverUser(row);
  return {
    id: row.id,
    name: row.name,
    email: row.email,
    zone: row.zone || "",
    role: observer ? "Admin" : row.role,
    accessLevel: observer ? OBSERVER_ACCESS_LEVEL : "",
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

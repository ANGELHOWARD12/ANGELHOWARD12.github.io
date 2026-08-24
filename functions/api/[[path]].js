const SESSION_COOKIE = "lg_session";
const SESSION_DAYS = 14;
const COORDINATOR_CODE_HASH = "e62163b1947feab8e4db70a99cffd5fb9c9f66d5e8901a4fb9775180ea780b71";
const MICROSOFT_STORAGE_PROVIDER = "onedrive";
const R2_STORAGE_PROVIDER = "r2";
const R2_SEGMENTED_STORAGE_PROVIDER = "r2-segmented";
const MAINTENANCE_INTERVAL_MS = 12 * 60 * 60 * 1000;
const STALE_MAINTENANCE_MS = 3 * 60 * 1000;
const ABANDONED_UPLOAD_TTL_MS = 10 * 60 * 1000;
const ABANDONED_MULTIPART_TTL_MS = 24 * 60 * 60 * 1000;
const DELIVERED_NOTIFICATION_TTL_MS = 14 * 24 * 60 * 60 * 1000;
const UNDELIVERED_NOTIFICATION_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const MAINTENANCE_SETTING_KEY = "storage_maintenance_structured_v1";
const LEGACY_EVIDENCE_MIGRATION_ENABLED = true;
const R2_MULTIPART_PART_BYTES = 5 * 1024 * 1024;
const WEEKLY_BACKUP_RETENTION = 12;
const SCHEMA_VERSION = "27-uppercase-users-1";
const OBSERVER_ACCESS_LEVEL = "observer";
const OBSERVER_EMAIL = "giuliana.parra@lgtask.local";
const PRIMARY_COORDINATOR_EMAIL = "pablo.ramos@lgtask.local";
const TEAM_TRAINING = "Training";
const TEAM_AUDIOVISUAL = "Audiovisuales";
const ORGANIZATION_USERS = [
  {
    id: "org-master-nykol-ruiz",
    name: "NYKOL RUIZ",
    email: "nykol.ruiz@lgtask.local",
    role: "Trainer",
    team: TEAM_TRAINING,
    jobTitle: "Master Trainer HS",
    memberType: "master",
    salt: "5mAD4vip7tnBdEorr_yyGw",
    hash: "2dvHhzW1XhdoqJ6q-tSoqawg9S3vHZuTu2TvHWOh3-Q"
  },
  {
    id: "org-master-ronald-chavez",
    name: "RONALD CHAVEZ",
    email: "ronald.chavez@lgtask.local",
    role: "Trainer",
    team: TEAM_TRAINING,
    jobTitle: "Master Trainer TV",
    memberType: "master",
    salt: "-JcPKjV6Dpos-bRGq4NkgQ",
    hash: "6vSvzRaKzrZCgw-fq8dfU9ONjOnpjwm3_8vMYbHBO08"
  },
  {
    id: "org-master-alejandro-cotrina",
    name: "ALEJANDRO COTRINA",
    email: "alejandro.cotrina@lgtask.local",
    role: "Coordinador",
    team: TEAM_AUDIOVISUAL,
    jobTitle: "Coordinador Audiovisual",
    memberType: "master",
    salt: "FD1wPgxZyy94gFvpsM_3YQ",
    hash: "7zlb3nk1N467QpagpEUZUeErwAgPbSMGsce38fgjUos"
  },
  {
    id: "org-av-ariana-perez",
    name: "ARIANA PEREZ",
    email: "ariana.perez@lgtask.local",
    role: "Trainer",
    team: TEAM_AUDIOVISUAL,
    jobTitle: "Creadora de Contenido",
    memberType: "audiovisual",
    salt: "FAv04x26nqTU8AHbbRJb3Q",
    hash: "9c1liZu0X1GdW0XPPNX_55Dj1qw5HZuIIX4VtjRbftw"
  },
  {
    id: "org-av-abel-barrantes",
    name: "ABEL BARRANTES",
    email: "abel.barrantes@lgtask.local",
    role: "Trainer",
    team: TEAM_AUDIOVISUAL,
    jobTitle: "Creador de Contenido",
    memberType: "audiovisual",
    salt: "AfnB1a4QbSujNCC63YaLCw",
    hash: "2wrchK-w7Eo6MR3c3uPcZ8VnIvwY1ygg-Rq3UyVE8VU"
  },
  {
    id: "org-av-fernando-bedrinana",
    name: "FERNANDO BEDRINANA",
    email: "fernando.bedrinana@lgtask.local",
    role: "Trainer",
    team: TEAM_AUDIOVISUAL,
    jobTitle: "Creador de Contenido",
    memberType: "audiovisual",
    salt: "yq4FgSz9-8FP9agCtsGQmw",
    hash: "hBoCgmrEp0lG70-LtB7JdnK3fZZ4ZtWO6lsANNz-T5o"
  },
  {
    id: "org-av-nathaly-fuentes",
    name: "NATHALY FUENTES",
    email: "nathaly.fuentes@lgtask.local",
    role: "Trainer",
    team: TEAM_AUDIOVISUAL,
    jobTitle: "Creadora de Contenido",
    memberType: "audiovisual",
    salt: "f3rT7wdXGFqP0UeMtlsuHA",
    hash: "98O4qYOFs6hHKX-4a40IqFinRUI-D1FVV4Xpz9b2zB8"
  }
];

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
    const url = new URL(request.url);
    const route = url.pathname.replace(/^\/(?:api|cloud)\/?/, "");

    if (request.method === "OPTIONS") return new Response(null, { status: 204 });
    if (!["GET", "HEAD"].includes(request.method) && !sameOrigin(request)) {
      return json({ ok: false, message: "Solicitud no permitida." }, 403);
    }

    if (route === "health" && request.method === "GET") {
      await runMaintenance(env.DB, env);
      return await healthStatus(env.DB, env);
    }
    if (route === "auth/register" && request.method === "POST") return await register(request, env.DB);
    if (route === "auth/login" && request.method === "POST") return await login(request, env.DB);
    if (route === "auth/logout" && request.method === "POST") return await logout(request, env.DB);
    if (route === "auth/recovery" && request.method === "POST") return await requestRecovery(request, env.DB);

    const session = await authenticate(request, env.DB);
    if (!session) return json({ ok: false, message: "Sesion no valida." }, 401);
    if ((route === "state" && request.method === "GET") || route === "storage/status") {
      scheduleMaintenance(env.DB, env, context);
    }
    if (
      isObserverUser(session.user) &&
      !["GET", "HEAD"].includes(request.method) &&
      !["notifications/subscribe", "notifications/unsubscribe"].includes(route)
    ) {
      return json({ ok: false, message: "La cuenta observadora es de solo lectura." }, 403);
    }

    if (route === "notifications/config" && request.method === "GET") return await notificationConfig(env.DB);
    if (route === "notifications/subscribe" && request.method === "POST") {
      return await subscribeNotifications(request, env.DB, session.user, context);
    }
    if (route === "notifications/unsubscribe" && request.method === "POST") {
      return await unsubscribeNotifications(request, env.DB, session.user);
    }
    if (route === "notifications/pending" && request.method === "GET") {
      return await pendingNotifications(env.DB, session.user);
    }
    if (route === "storage/status" && request.method === "GET") return await storageStatus(env.DB, session.user, env);
    if (route === "storage/backup" && request.method === "POST") return await createStorageBackup(env.DB, session.user, env);
    if (route === "evidence/upload/r2" && request.method === "POST") {
      return await uploadEvidenceDirectToR2(request, env.DB, session.user, env);
    }
    if (route === "evidence/upload/r2/init" && request.method === "POST") {
      return await initR2MultipartUpload(request, env.DB, session.user, env);
    }
    if (route === "evidence/upload/r2/part" && request.method === "POST") {
      return await uploadR2MultipartPart(request, env.DB, session.user, env);
    }
    if (route === "evidence/upload/r2/complete" && request.method === "POST") {
      return await completeR2MultipartUpload(request, env.DB, session.user, env);
    }
    if (route === "evidence/upload" && request.method === "POST") return await uploadEvidence(request, env.DB, session.user, env);
    if (route === "evidence/upload/init" && request.method === "POST") return await initEvidenceUpload(request, env.DB, session.user);
    if (route === "evidence/upload/chunk" && request.method === "POST") return await uploadEvidenceChunk(request, env.DB, session.user);
    if (route === "evidence/upload/complete" && request.method === "POST") return await completeEvidenceUpload(request, env.DB, session.user, env);
    const evidenceFileMatch = route.match(/^evidence\/([^/]+)\/(?:file|photo)$/);
    if (evidenceFileMatch && request.method === "GET") return await evidenceFile(env.DB, session.user, evidenceFileMatch[1], env);
    if (route === "tasks/evidence" && request.method === "POST") return await submitTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/create" && request.method === "POST") return await createTask(request, env.DB, session.user, context);
    if (route === "tasks/master-update" && request.method === "POST") return await updateMasterTask(request, env.DB, session.user);
    if (route === "tasks/evidence-authorize" && request.method === "POST") return await authorizeLateTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/review" && request.method === "POST") return await reviewTaskEvidence(request, env.DB, session.user, context);
    if (route === "tasks/delete" && request.method === "POST") return await deleteTaskAndArchive(request, env.DB, session.user, context, env);
    if (route === "schedule/overtime" && request.method === "POST") return await saveOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "schedule/overtime-request" && request.method === "POST") return await requestOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "schedule/overtime-review" && request.method === "POST") return await reviewOvertimeSchedule(request, env.DB, session.user, context);
    if (route === "state" && request.method === "GET") return await getState(env.DB, session.user, context);
    if (route === "state" && request.method === "PUT") return await putState(request, env.DB, session.user, context, env);
    if (route === "settings/late-evidence" && request.method === "POST") {
      return await setLateEvidencePolicy(request, env.DB, session.user, context);
    }
    if (route === "info/updates" && request.method === "POST") return await saveInfoUpdate(request, env.DB, session.user);
    if (route === "admin/users" && request.method === "POST") return await createUser(request, env.DB, session.user);
    if (route === "admin/reset-password" && request.method === "POST") return await resetPassword(request, env.DB, session.user);

    return json({ ok: false, message: "Ruta no encontrada." }, 404);
  } catch (error) {
    console.error("LG Task API", error);
    const status = Number(error?.status || 500);
    return json(
      {
        ok: false,
        message: status < 500 ? clean(error?.message) || "La solicitud no es valida." : "No se pudo completar la operacion en la nube."
      },
      status >= 400 && status < 600 ? status : 500
    );
  }
}

async function healthStatus(db, env) {
  const [pendingResult, failedResult, maintenanceResult, backupResult] = await db.batch([
    db.prepare(
      `SELECT COUNT(*) AS count FROM evidence_files files
       WHERE NOT EXISTS (SELECT 1 FROM evidence_storage storage WHERE storage.file_id = files.id)
         AND (LENGTH(files.photo_base64) > 0 OR EXISTS (
           SELECT 1 FROM evidence_file_chunks chunks WHERE chunks.file_id = files.id
         ))`
    ),
    db.prepare("SELECT COUNT(*) AS count FROM app_settings WHERE key LIKE 'storage_migration_failed:%'"),
    db.prepare("SELECT value, updated_at FROM app_settings WHERE key = ?").bind(MAINTENANCE_SETTING_KEY),
    db.prepare("SELECT updated_at FROM app_settings WHERE key LIKE 'weekly_backup:%' ORDER BY updated_at DESC LIMIT 1")
  ]);
  return json({
    ok: true,
    version: "31-task-hub-ui1",
    schema: SCHEMA_VERSION,
    r2: r2StorageEnabled(env),
    migration: {
      pendingFiles: Number(pendingResult?.results?.[0]?.count || 0),
      failedFiles: Number(failedResult?.results?.[0]?.count || 0),
      state: clean(maintenanceResult?.results?.[0]?.value) || "pending",
      updatedAt: Number(maintenanceResult?.results?.[0]?.updated_at || 0)
    },
    backup: {
      ready: Number(backupResult?.results?.[0]?.updated_at || 0) > 0,
      updatedAt: Number(backupResult?.results?.[0]?.updated_at || 0)
    }
  });
}

async function ensureSchema(db) {
  if (schemaReady) return;
  try {
    const marker = await db.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'").first();
    if (clean(marker?.value) === SCHEMA_VERSION) {
      schemaReady = true;
      return;
    }
  } catch (error) {
    if (!String(error?.message || error).toLowerCase().includes("no such table")) throw error;
  }
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
    db.prepare(`CREATE TABLE IF NOT EXISTS task_records (
      id TEXT PRIMARY KEY,
      owner_id TEXT NOT NULL DEFAULT '',
      due_date TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT '',
      archived INTEGER NOT NULL DEFAULT 0,
      created_at INTEGER NOT NULL DEFAULT 0,
      data TEXT NOT NULL,
      updated_at INTEGER NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS task_history_records (
      task_id TEXT NOT NULL,
      entry_index INTEGER NOT NULL,
      data TEXT NOT NULL,
      PRIMARY KEY (task_id, entry_index),
      FOREIGN KEY (task_id) REFERENCES task_records(id) ON DELETE CASCADE
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS r2_multipart_uploads (
      id TEXT PRIMARY KEY,
      task_id TEXT NOT NULL,
      owner_id TEXT NOT NULL,
      submitted_by_id TEXT NOT NULL,
      client_key TEXT NOT NULL,
      file_name TEXT NOT NULL,
      mime_type TEXT NOT NULL,
      size_bytes INTEGER NOT NULL,
      object_key TEXT NOT NULL,
      parent_path TEXT NOT NULL DEFAULT '',
      r2_upload_id TEXT NOT NULL,
      parts_json TEXT NOT NULL DEFAULT '[]',
      created_at INTEGER NOT NULL,
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
    db.prepare("CREATE INDEX IF NOT EXISTS idx_user_notifications_pending ON user_notifications(user_id, delivered_at)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_task_records_owner_date ON task_records(owner_id, due_date)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_task_records_status_date ON task_records(status, due_date)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_task_records_archive_created ON task_records(archived, created_at)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_task_history_task ON task_history_records(task_id, entry_index)"),
    db.prepare("CREATE UNIQUE INDEX IF NOT EXISTS idx_r2_multipart_client ON r2_multipart_uploads(submitted_by_id, task_id, client_key)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_r2_multipart_created ON r2_multipart_uploads(created_at)")
  ]);
  const userColumns = await db.prepare("PRAGMA table_info(users)").all();
  if (!(userColumns.results || []).some((column) => column.name === "access_level")) {
    try {
      await db.prepare("ALTER TABLE users ADD COLUMN access_level TEXT NOT NULL DEFAULT ''").run();
    } catch (error) {
      if (!String(error?.message || error).toLowerCase().includes("duplicate column")) throw error;
    }
  }
  for (const [column, definition] of [
    ["team", "TEXT NOT NULL DEFAULT 'Training'"],
    ["job_title", "TEXT NOT NULL DEFAULT ''"],
    ["member_type", "TEXT NOT NULL DEFAULT 'trainer'"]
  ]) {
    if (!(userColumns.results || []).some((item) => item.name === column)) {
      try {
        await db.prepare(`ALTER TABLE users ADD COLUMN ${column} ${definition}`).run();
      } catch (error) {
        if (!String(error?.message || error).toLowerCase().includes("duplicate column")) throw error;
      }
    }
  }
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_users_team_status ON users(team, status, role)").run();
  await db
    .prepare("UPDATE users SET access_level = ?, zone = 'Administracion general', team = 'Todos', job_title = 'Admin', member_type = 'admin' WHERE email = ?")
    .bind(OBSERVER_ACCESS_LEVEL, OBSERVER_EMAIL)
    .run();
  await db
    .prepare("UPDATE users SET team = ?, job_title = 'Coordinador de Entrenamiento', member_type = 'master' WHERE email = ?")
    .bind(TEAM_TRAINING, PRIMARY_COORDINATOR_EMAIL)
    .run();
  await db.prepare("UPDATE users SET name = UPPER(TRIM(name)) WHERE name <> UPPER(TRIM(name))").run();
  await db.batch(
    ORGANIZATION_USERS.map((profile) =>
      db
        .prepare(
          `INSERT OR IGNORE INTO users
           (id, name, email, zone, role, status, password_hash, password_salt, created_at, access_level, team, job_title, member_type)
           VALUES (?, ?, ?, ?, ?, 'Activo', ?, ?, ?, '', ?, ?, ?)`
        )
        .bind(
          profile.id,
          profile.name,
          profile.email,
          profile.team,
          profile.role,
          profile.hash,
          profile.salt,
          Date.now(),
          profile.team,
          profile.jobTitle,
          profile.memberType
        )
    )
  );
  await db.batch(
    ORGANIZATION_USERS.map((profile) =>
      db
        .prepare("UPDATE users SET team = ?, job_title = ?, member_type = ? WHERE email = ?")
        .bind(profile.team, profile.jobTitle, profile.memberType, profile.email)
    )
  );
  await migrateStructuredTasks(db);
  await db
    .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES ('schema_version', ?, ?)")
    .bind(SCHEMA_VERSION, Date.now())
    .run();
  schemaReady = true;
}

async function migrateStructuredTasks(db) {
  const source = await db
    .prepare(
      `SELECT
         COALESCE(json_array_length(data, '$.tasks'), 0) AS active_count,
         COALESCE(json_array_length(data, '$.deletedTasks'), 0) AS archived_count
       FROM app_data WHERE id = 1`
    )
    .first();
  const expected = Number(source?.active_count || 0) + Number(source?.archived_count || 0);
  if (expected > 0) {
    const now = Date.now();
    await db.batch([
      db.prepare(
        `INSERT OR REPLACE INTO task_records
         (id, owner_id, due_date, status, archived, created_at, data, updated_at)
         SELECT json_extract(task.value, '$.id'),
                COALESCE(json_extract(task.value, '$.ownerId'), ''),
                COALESCE(json_extract(task.value, '$.dueDate'), ''),
                COALESCE(json_extract(task.value, '$.status'), ''),
                0,
                COALESCE(json_extract(task.value, '$.createdAt'), 0),
                json_remove(task.value, '$.history'), ?
         FROM app_data data, json_each(data.data, '$.tasks') task
         WHERE data.id = 1 AND COALESCE(json_extract(task.value, '$.id'), '') <> ''`
      ).bind(now),
      db.prepare(
        `INSERT OR REPLACE INTO task_records
         (id, owner_id, due_date, status, archived, created_at, data, updated_at)
         SELECT json_extract(task.value, '$.id'),
                COALESCE(json_extract(task.value, '$.ownerId'), ''),
                COALESCE(json_extract(task.value, '$.dueDate'), ''),
                COALESCE(json_extract(task.value, '$.status'), ''),
                1,
                COALESCE(json_extract(task.value, '$.deletedAt'), json_extract(task.value, '$.createdAt'), 0),
                json_remove(task.value, '$.history'), ?
         FROM app_data data, json_each(data.data, '$.deletedTasks') task
         WHERE data.id = 1 AND COALESCE(json_extract(task.value, '$.id'), '') <> ''`
      ).bind(now),
      db.prepare(
        `INSERT OR REPLACE INTO task_history_records (task_id, entry_index, data)
         SELECT json_extract(task.value, '$.id'), CAST(history.key AS INTEGER), history.value
         FROM app_data data,
              json_each(data.data, '$.tasks') task,
              json_each(task.value, '$.history') history
         WHERE data.id = 1 AND COALESCE(json_extract(task.value, '$.id'), '') <> ''`
      ),
      db.prepare(
        `INSERT OR REPLACE INTO task_history_records (task_id, entry_index, data)
         SELECT json_extract(task.value, '$.id'), CAST(history.key AS INTEGER), history.value
         FROM app_data data,
              json_each(data.data, '$.deletedTasks') task,
              json_each(task.value, '$.history') history
         WHERE data.id = 1 AND COALESCE(json_extract(task.value, '$.id'), '') <> ''`
      )
    ]);
  }

  const stored = await db.prepare("SELECT COUNT(*) AS count FROM task_records").first();
  if (Number(stored?.count || 0) < expected) {
    throw new Error(`Structured task migration incomplete: expected ${expected}, stored ${Number(stored?.count || 0)}`);
  }
  await db
    .prepare("UPDATE app_data SET data = json_remove(data, '$.tasks', '$.deletedTasks'), updated_at = ? WHERE id = 1")
    .bind(Date.now())
    .run();
}

function scheduleMaintenance(db, env, context) {
  context.waitUntil(
    runMaintenance(db, env).catch((error) => {
      console.error("Task Hub maintenance", error);
    })
  );
}

async function runMaintenance(db, env) {
  const now = Date.now();
  const claim = await db
    .prepare(
      `INSERT INTO app_settings (key, value, updated_at) VALUES (?, 'running', ?)
       ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
       WHERE app_settings.updated_at < ?
          OR (app_settings.value = 'running' AND app_settings.updated_at < ?)`
    )
    .bind(MAINTENANCE_SETTING_KEY, now, now - MAINTENANCE_INTERVAL_MS, now - STALE_MAINTENANCE_MS)
    .run();
  if (Number(claim?.meta?.changes || 0) === 0) return;

  const abandonedCutoff = now - ABANDONED_UPLOAD_TTL_MS;
  const notificationCutoff = now - DELIVERED_NOTIFICATION_TTL_MS;
  const undeliveredNotificationCutoff = now - UNDELIVERED_NOTIFICATION_TTL_MS;
  const unreferencedFileIds = `
    SELECT ef.id FROM evidence_files ef
    WHERE ef.created_at < ?
      AND NOT EXISTS (
        SELECT 1 FROM task_records task WHERE task.id = ef.task_id
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
  await cleanupAbandonedMultipartUploads(db, env, now - ABANDONED_MULTIPART_TTL_MS);

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

  const migration = LEGACY_EVIDENCE_MIGRATION_ENABLED
    ? await migrateOneLegacyEvidence(db, env)
    : { migrated: false, failed: false };
  if (migration.migrated || migration.failed || migration.progressed) {
    await db
      .prepare("UPDATE app_settings SET value = ?, updated_at = 0 WHERE key = ?")
      .bind(migration.failed ? "migration-retry" : "migrating", MAINTENANCE_SETTING_KEY)
      .run();
    return;
  }

  await ensureWeeklyBackup(db, env).catch((error) => {
    console.error("Weekly R2 backup", error);
  });

  await db
    .prepare("UPDATE app_settings SET value = 'complete', updated_at = ? WHERE key = ?")
    .bind(Date.now(), MAINTENANCE_SETTING_KEY)
    .run();
}

async function migrateOneLegacyEvidence(db, env) {
  if (!r2StorageEnabled(env)) return { migrated: false, failed: false };
  const candidates = await db
    .prepare(
      `SELECT files.id, files.task_id, files.owner_id, files.submitted_by_id,
              files.file_name, files.mime_type, files.created_at,
              LENGTH(files.photo_base64) AS photo_base64_length,
              SUBSTR(files.photo_base64, -2) AS photo_base64_tail
       FROM evidence_files files
       WHERE NOT EXISTS (SELECT 1 FROM evidence_storage storage WHERE storage.file_id = files.id)
         AND (LENGTH(files.photo_base64) > 0 OR EXISTS (
           SELECT 1 FROM evidence_file_chunks chunks WHERE chunks.file_id = files.id
         ))
         AND EXISTS (
           SELECT 1 FROM task_records task WHERE task.id = files.task_id AND task.archived = 0
         )
         AND NOT EXISTS (
           SELECT 1 FROM app_settings setting WHERE setting.key = 'storage_migration_failed:' || files.id
       )
       ORDER BY files.created_at ASC
       LIMIT 1`
    )
    .all();
  if (!candidates.results?.length) return { migrated: false, failed: false };

  const data = await loadData(db);
  const results = await Promise.all(
    candidates.results.map(async (candidate) => {
      try {
        const task = data.tasks.find((item) => clean(item.id) === clean(candidate.task_id));
        if (!task) throw new Error("Referenced task not found");
        const file = { id: candidate.id, fileName: candidate.file_name, mimeType: candidate.mime_type };
        let completed = true;
        if (Number(candidate.photo_base64_length || 0) > 0) {
          completed = await migrateBase64EvidenceSegmentStep(env, db, task, file, candidate);
        } else {
          completed = await migrateChunkedEvidenceSegmentStep(env, db, task, file);
        }
        if (!completed) return { migrated: false, failed: false, progressed: true };
        await db.batch([
          db.prepare("UPDATE evidence_files SET photo_base64 = '' WHERE id = ?").bind(candidate.id),
          db.prepare("DELETE FROM evidence_file_chunks WHERE file_id = ?").bind(candidate.id),
          db.prepare("DELETE FROM app_settings WHERE key = ?").bind(`storage_migration_progress:${candidate.id}`)
        ]);
        return { migrated: true, failed: false, progressed: true };
      } catch (error) {
        console.error("Legacy evidence migration", candidate.id, error);
        if (!String(error?.message || error).startsWith("PERMANENT:")) {
          return { migrated: false, failed: false, progressed: true };
        }
        await db
          .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
          .bind(`storage_migration_failed:${candidate.id}`, clean(error?.message || error).slice(0, 300), Date.now())
          .run();
        return { migrated: false, failed: true, progressed: false };
      }
    })
  );
  return {
    migrated: results.some((result) => result.migrated),
    failed: results.some((result) => result.failed),
    progressed: results.some((result) => result.progressed)
  };
}

async function legacySegmentProgress(env, db, task, file, source, totalBytes, totalSegments) {
  const settingKey = `storage_migration_progress:${file.id}`;
  const saved = await db.prepare("SELECT value FROM app_settings WHERE key = ?").bind(settingKey).first();
  let progress = null;
  if (saved?.value) {
    progress = JSON.parse(saved.value);
    if (progress.mode !== "segments") {
      if (progress.objectKey && progress.r2UploadId) {
        await env.EVIDENCE_BUCKET.resumeMultipartUpload(progress.objectKey, progress.r2UploadId).abort().catch(() => {});
      }
      await db.prepare("DELETE FROM app_settings WHERE key = ?").bind(settingKey).run();
      progress = null;
    }
  }
  if (!progress) {
    const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
    progress = {
      mode: "segments",
      source,
      objectKey,
      manifestKey: `${objectKey}.manifest.json`,
      parentPath,
      totalBytes,
      totalSegments,
      segments: []
    };
    await db
      .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
      .bind(settingKey, JSON.stringify(progress), Date.now())
      .run();
  }
  if (
    progress.source !== source ||
    Number(progress.totalBytes) !== totalBytes ||
    Number(progress.totalSegments) !== totalSegments
  ) {
    throw new Error("PERMANENT: Legacy evidence changed during segmented migration");
  }
  return { settingKey, progress };
}

async function saveLegacyR2Segment(env, db, file, settingKey, progress, bytes) {
  const segmentIndex = progress.segments.length;
  const segmentKey = `${progress.objectKey}.segments/${String(segmentIndex).padStart(4, "0")}`;
  await env.EVIDENCE_BUCKET.put(segmentKey, bytes, {
    httpMetadata: { contentType: "application/octet-stream" },
    customMetadata: { fileId: clean(file.id), segment: String(segmentIndex) }
  });
  progress.segments.push({ key: segmentKey, size: bytes.byteLength });
  await db
    .prepare("UPDATE app_settings SET value = ?, updated_at = ? WHERE key = ?")
    .bind(JSON.stringify(progress), Date.now(), settingKey)
    .run();
}

async function completeLegacyR2Segments(env, db, file, progress) {
  const storedBytes = progress.segments.reduce((total, segment) => total + Number(segment.size || 0), 0);
  if (storedBytes !== Number(progress.totalBytes)) throw new Error("PERMANENT: Segmented R2 size mismatch");
  const manifest = {
    version: 1,
    fileId: clean(file.id),
    mimeType: clean(file.mimeType),
    size: Number(progress.totalBytes),
    segments: progress.segments
  };
  await env.EVIDENCE_BUCKET.put(progress.manifestKey, JSON.stringify(manifest), {
    httpMetadata: { contentType: "application/json" },
    customMetadata: { type: "evidence-manifest", fileId: clean(file.id) }
  });
  await db
    .prepare(
      `INSERT OR REPLACE INTO evidence_storage
       (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
       VALUES (?, ?, ?, ?, ?, ?, '', ?)`
    )
    .bind(
      file.id,
      R2_SEGMENTED_STORAGE_PROVIDER,
      "EVIDENCE_BUCKET",
      progress.manifestKey,
      progress.parentPath,
      progress.totalBytes,
      Date.now()
    )
    .run();
  return true;
}

async function migrateBase64EvidenceSegmentStep(env, db, task, file, candidate) {
  const base64Length = Number(candidate.photo_base64_length || 0);
  const tail = String(candidate.photo_base64_tail || "");
  const padding = tail.endsWith("==") ? 2 : tail.endsWith("=") ? 1 : 0;
  const totalBytes = Math.floor((base64Length * 3) / 4) - padding;
  if (base64Length < 4 || totalBytes < 1 || totalBytes > MAX_FILE_BYTES) {
    throw new Error("PERMANENT: Invalid legacy Base64 file size");
  }
  const segmentBase64Chars = 900_000;
  const totalSegments = Math.ceil(base64Length / segmentBase64Chars);
  const { settingKey, progress } = await legacySegmentProgress(
    env,
    db,
    task,
    file,
    "base64-segments",
    totalBytes,
    totalSegments
  );
  const segmentIndex = progress.segments.length;
  if (segmentIndex < totalSegments) {
    const row = await db
      .prepare("SELECT SUBSTR(photo_base64, ?, ?) AS part_base64 FROM evidence_files WHERE id = ?")
      .bind(segmentIndex * segmentBase64Chars + 1, segmentBase64Chars, file.id)
      .first();
    const partBase64 = String(row?.part_base64 || "");
    if (!partBase64 || (segmentIndex < totalSegments - 1 && partBase64.length !== segmentBase64Chars)) {
      throw new Error("PERMANENT: Legacy Base64 segment is incomplete");
    }
    await saveLegacyR2Segment(env, db, file, settingKey, progress, decodeEvidenceBase64(partBase64));
  }
  if (progress.segments.length < totalSegments) return false;
  return completeLegacyR2Segments(env, db, file, progress);
}

async function migrateChunkedEvidenceSegmentStep(env, db, task, file) {
  const chunks = await db
    .prepare(
      `SELECT chunk_index, LENGTH(chunk_base64) AS base64_length,
              SUBSTR(chunk_base64, -2) AS base64_tail
       FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC`
    )
    .bind(file.id)
    .all();
  const chunkRows = chunks.results || [];
  if (!chunkRows.length || chunkRows.some((chunk, index) => Number(chunk.chunk_index) !== index)) {
    throw new Error("PERMANENT: Incomplete legacy chunks");
  }
  const decodedLength = (chunk) => {
    const tail = String(chunk.base64_tail || "");
    const padding = tail.endsWith("==") ? 2 : tail.endsWith("=") ? 1 : 0;
    return Math.floor((Number(chunk.base64_length || 0) * 3) / 4) - padding;
  };
  const totalBytes = chunkRows.reduce((total, chunk) => total + decodedLength(chunk), 0);
  if (totalBytes < 1 || totalBytes > MAX_FILE_BYTES) throw new Error("PERMANENT: Invalid legacy chunk size");
  const { settingKey, progress } = await legacySegmentProgress(
    env,
    db,
    task,
    file,
    "chunk-segments",
    totalBytes,
    chunkRows.length
  );
  const segmentIndex = progress.segments.length;
  if (segmentIndex < chunkRows.length) {
    const row = await db
      .prepare("SELECT chunk_base64 FROM evidence_file_chunks WHERE file_id = ? AND chunk_index = ?")
      .bind(file.id, segmentIndex)
      .first();
    if (!row?.chunk_base64) throw new Error("PERMANENT: Legacy chunk segment is incomplete");
    await saveLegacyR2Segment(env, db, file, settingKey, progress, decodeEvidenceBase64(row.chunk_base64));
  }
  if (progress.segments.length < chunkRows.length) return false;
  return completeLegacyR2Segments(env, db, file, progress);
}

async function migrateBase64EvidenceStep(env, db, task, file, candidate) {
  const base64Length = Number(candidate.photo_base64_length || 0);
  const tail = String(candidate.photo_base64_tail || "");
  const padding = tail.endsWith("==") ? 2 : tail.endsWith("=") ? 1 : 0;
  const totalBytes = Math.floor((base64Length * 3) / 4) - padding;
  if (base64Length < 4 || totalBytes < 1 || totalBytes > MAX_FILE_BYTES) {
    throw new Error("PERMANENT: Invalid legacy Base64 file size");
  }
  const partBase64Chars = 7_200_000;
  const totalGroups = Math.ceil(base64Length / partBase64Chars);
  const settingKey = `storage_migration_progress:${file.id}`;
  const saved = await db.prepare("SELECT value FROM app_settings WHERE key = ?").bind(settingKey).first();
  let progress;
  if (saved?.value) {
    progress = JSON.parse(saved.value);
  } else {
    const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
    const upload = await env.EVIDENCE_BUCKET.createMultipartUpload(objectKey, {
      httpMetadata: { contentType: file.mimeType },
      customMetadata: { fileId: clean(file.id), taskId: clean(task.id), ownerId: clean(task.ownerId) }
    });
    progress = {
      source: "base64",
      objectKey,
      parentPath,
      r2UploadId: upload.uploadId,
      totalBytes,
      totalGroups,
      parts: []
    };
    await db
      .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
      .bind(settingKey, JSON.stringify(progress), Date.now())
      .run();
  }
  if (
    clean(progress.source) !== "base64" ||
    Number(progress.totalBytes) !== totalBytes ||
    Number(progress.totalGroups) !== totalGroups
  ) {
    throw new Error("PERMANENT: Legacy Base64 changed during migration");
  }

  const groupIndex = progress.parts.length;
  if (groupIndex < totalGroups) {
    const offset = groupIndex * partBase64Chars;
    const row = await db
      .prepare("SELECT SUBSTR(photo_base64, ?, ?) AS part_base64 FROM evidence_files WHERE id = ?")
      .bind(offset + 1, partBase64Chars, file.id)
      .first();
    const partBase64 = String(row?.part_base64 || "");
    if (!partBase64 || (groupIndex < totalGroups - 1 && partBase64.length !== partBase64Chars)) {
      throw new Error("PERMANENT: Legacy Base64 part is incomplete");
    }
    const bytes = decodeEvidenceBase64(partBase64);
    if (groupIndex < totalGroups - 1 && bytes.byteLength < R2_MULTIPART_PART_BYTES) {
      throw new Error("PERMANENT: Legacy Base64 multipart part is too small");
    }
    const uploaded = await env.EVIDENCE_BUCKET
      .resumeMultipartUpload(progress.objectKey, progress.r2UploadId)
      .uploadPart(groupIndex + 1, bytes);
    progress.parts.push({ partNumber: groupIndex + 1, etag: uploaded.etag });
    await db
      .prepare("UPDATE app_settings SET value = ?, updated_at = ? WHERE key = ?")
      .bind(JSON.stringify(progress), Date.now(), settingKey)
      .run();
  }
  if (progress.parts.length < totalGroups) return false;

  let storedObject;
  try {
    storedObject = await env.EVIDENCE_BUCKET
      .resumeMultipartUpload(progress.objectKey, progress.r2UploadId)
      .complete(progress.parts);
  } catch (error) {
    storedObject = await env.EVIDENCE_BUCKET.head(progress.objectKey).catch(() => null);
    if (!storedObject || Number(storedObject.size || 0) !== totalBytes) throw error;
  }
  if (Number(storedObject?.size || 0) !== totalBytes) throw new Error("PERMANENT: Incomplete Base64 R2 migration");
  await db
    .prepare(
      `INSERT OR REPLACE INTO evidence_storage
       (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
       VALUES (?, ?, ?, ?, ?, ?, '', ?)`
    )
    .bind(file.id, R2_STORAGE_PROVIDER, "EVIDENCE_BUCKET", progress.objectKey, progress.parentPath, totalBytes, Date.now())
    .run();
  return true;
}

async function migrateChunkedEvidenceStep(env, db, task, file) {
  const chunks = await db
    .prepare(
      `SELECT chunk_index, LENGTH(chunk_base64) AS base64_length,
              SUBSTR(chunk_base64, -2) AS base64_tail
       FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC`
    )
    .bind(file.id)
    .all();
  const chunkRows = chunks.results || [];
  if (!chunkRows.length || chunkRows.some((chunk, index) => Number(chunk.chunk_index) !== index)) {
    throw new Error("PERMANENT: Incomplete legacy chunks");
  }
  const decodedLength = (chunk) => {
    const tail = String(chunk.base64_tail || "");
    const padding = tail.endsWith("==") ? 2 : tail.endsWith("=") ? 1 : 0;
    return Math.floor((Number(chunk.base64_length || 0) * 3) / 4) - padding;
  };
  const totalBytes = chunkRows.reduce((total, chunk) => total + decodedLength(chunk), 0);
  if (totalBytes < 1 || totalBytes > MAX_FILE_BYTES) throw new Error("PERMANENT: Invalid legacy file size");

  const groups = [];
  let currentGroup = [];
  let currentBytes = 0;
  for (const chunk of chunkRows) {
    currentGroup.push(chunk);
    currentBytes += decodedLength(chunk);
    if (currentBytes >= R2_MULTIPART_PART_BYTES) {
      groups.push(currentGroup);
      currentGroup = [];
      currentBytes = 0;
    }
  }
  if (currentGroup.length) groups.push(currentGroup);

  const settingKey = `storage_migration_progress:${file.id}`;
  const saved = await db.prepare("SELECT value FROM app_settings WHERE key = ?").bind(settingKey).first();
  let progress;
  if (saved?.value) {
    progress = JSON.parse(saved.value);
  } else {
    const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
    const upload = await env.EVIDENCE_BUCKET.createMultipartUpload(objectKey, {
      httpMetadata: { contentType: file.mimeType },
      customMetadata: { fileId: clean(file.id), taskId: clean(task.id), ownerId: clean(task.ownerId) }
    });
    progress = {
      objectKey,
      parentPath,
      r2UploadId: upload.uploadId,
      totalBytes,
      totalGroups: groups.length,
      parts: []
    };
    await db
      .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
      .bind(settingKey, JSON.stringify(progress), Date.now())
      .run();
  }
  if (Number(progress.totalBytes) !== totalBytes || Number(progress.totalGroups) !== groups.length) {
    throw new Error("PERMANENT: Legacy chunks changed during migration");
  }

  const groupIndex = progress.parts.length;
  if (groupIndex < groups.length) {
    const group = groups[groupIndex];
    const firstIndex = Number(group[0].chunk_index);
    const lastIndex = Number(group[group.length - 1].chunk_index);
    const rows = await db
      .prepare(
        `SELECT chunk_base64 FROM evidence_file_chunks
         WHERE file_id = ? AND chunk_index BETWEEN ? AND ? ORDER BY chunk_index ASC`
      )
      .bind(file.id, firstIndex, lastIndex)
      .all();
    if ((rows.results || []).length !== group.length) throw new Error("PERMANENT: Legacy chunks changed during migration");
    const partBase64 = (rows.results || []).map((row) => String(row.chunk_base64 || "")).join("");
    const bytes = decodeEvidenceBase64(partBase64);
    if (groupIndex < groups.length - 1 && bytes.byteLength < R2_MULTIPART_PART_BYTES) {
      throw new Error("PERMANENT: Legacy multipart part is too small");
    }
    const uploaded = await env.EVIDENCE_BUCKET
      .resumeMultipartUpload(progress.objectKey, progress.r2UploadId)
      .uploadPart(groupIndex + 1, bytes);
    progress.parts.push({ partNumber: groupIndex + 1, etag: uploaded.etag });
    await db
      .prepare("UPDATE app_settings SET value = ?, updated_at = ? WHERE key = ?")
      .bind(JSON.stringify(progress), Date.now(), settingKey)
      .run();
  }
  if (progress.parts.length < groups.length) return false;

  let storedObject;
  try {
    storedObject = await env.EVIDENCE_BUCKET
      .resumeMultipartUpload(progress.objectKey, progress.r2UploadId)
      .complete(progress.parts);
  } catch (error) {
    storedObject = await env.EVIDENCE_BUCKET.head(progress.objectKey).catch(() => null);
    if (!storedObject || Number(storedObject.size || 0) !== totalBytes) throw error;
  }
  if (Number(storedObject?.size || 0) !== totalBytes) throw new Error("PERMANENT: Incomplete R2 multipart migration");
  await db
    .prepare(
      `INSERT OR REPLACE INTO evidence_storage
       (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
       VALUES (?, ?, ?, ?, ?, ?, '', ?)`
    )
    .bind(file.id, R2_STORAGE_PROVIDER, "EVIDENCE_BUCKET", progress.objectKey, progress.parentPath, totalBytes, Date.now())
    .run();
  return true;
}

async function cleanupAbandonedMultipartUploads(db, env, cutoff) {
  const rows = await db
    .prepare("SELECT id, object_key, r2_upload_id FROM r2_multipart_uploads WHERE updated_at < ? LIMIT 50")
    .bind(cutoff)
    .all();
  for (const row of rows.results || []) {
    if (r2StorageEnabled(env)) {
      await env.EVIDENCE_BUCKET.resumeMultipartUpload(row.object_key, row.r2_upload_id).abort().catch(() => {});
    }
    await db.prepare("DELETE FROM r2_multipart_uploads WHERE id = ?").bind(row.id).run();
  }
}

function isoWeekId(dateValue) {
  const date = new Date(`${dateValue}T12:00:00Z`);
  const weekday = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - weekday);
  const year = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(year, 0, 1));
  const week = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${year}-W${String(week).padStart(2, "0")}`;
}

async function ensureWeeklyBackup(db, env, { force = false } = {}) {
  if (!r2StorageEnabled(env)) throw new Error("R2 storage is not configured");
  const generatedAt = Date.now();
  const dateValue = dateIsoInLima(generatedAt);
  const weekId = isoWeekId(dateValue);
  const markerKey = `weekly_backup:${weekId}`;
  const existing = await db.prepare("SELECT value FROM app_settings WHERE key = ?").bind(markerKey).first();
  if (existing && !force) {
    try {
      return JSON.parse(existing.value);
    } catch {
      return { weekId, objectKey: clean(existing.value), generatedAt: 0 };
    }
  }

  const data = await loadData(db);
  const [usersResult, filesResult] = await db.batch([
    db.prepare(
      `SELECT id, name, email, zone, role, status, access_level, team, job_title, member_type, created_at
       FROM users ORDER BY created_at ASC`
    ),
    db.prepare(
      `SELECT files.id, files.task_id, files.owner_id, files.submitted_by_id,
              files.file_name, files.mime_type, files.created_at,
              storage.provider, storage.drive_item_id AS object_key,
              storage.parent_path, storage.size_bytes, storage.sha256
       FROM evidence_files files
       LEFT JOIN evidence_storage storage ON storage.file_id = files.id
       ORDER BY files.created_at ASC`
    )
  ]);
  const objectKey = `backups/${weekId.slice(0, 4)}/${weekId}/lgtask-${dateValue}-${generatedAt}.json`;
  const backup = {
    product: "Task Hub",
    schemaVersion: SCHEMA_VERSION,
    weekId,
    generatedAt,
    data,
    users: usersResult.results || [],
    evidenceFiles: filesResult.results || []
  };
  const body = JSON.stringify(backup);
  await env.EVIDENCE_BUCKET.put(objectKey, body, {
    httpMetadata: { contentType: "application/json" },
    customMetadata: { type: "weekly-backup", weekId, generatedAt: String(generatedAt) }
  });
  const result = { weekId, objectKey, generatedAt, bytes: new TextEncoder().encode(body).byteLength };
  await db
    .prepare("INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)")
    .bind(markerKey, JSON.stringify(result), generatedAt)
    .run();

  const listed = await env.EVIDENCE_BUCKET.list({ prefix: "backups/", limit: 1000 });
  const oldBackups = [...listed.objects]
    .sort((left, right) => new Date(right.uploaded || 0) - new Date(left.uploaded || 0))
    .slice(WEEKLY_BACKUP_RETENTION)
    .map((item) => item.key);
  if (oldBackups.length) await env.EVIDENCE_BUCKET.delete(oldBackups);
  return result;
}

async function createStorageBackup(db, user, env) {
  if (!isPrimaryCoordinatorUser(user)) {
    return json({ ok: false, message: "Solo Pablo puede generar respaldos manuales." }, 403);
  }
  try {
    const backup = await ensureWeeklyBackup(db, env, { force: true });
    return json({ ok: true, backup });
  } catch (error) {
    console.error("Manual R2 backup", error);
    return json({ ok: false, message: "No se pudo crear el respaldo en R2." }, 503);
  }
}

async function storageStatus(db, user, env) {
  if (!isPrimaryCoordinatorUser(user) && !isObserverUser(user)) {
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
         (SELECT COUNT(*) FROM task_records) AS task_rows,
         (SELECT COALESCE(SUM(LENGTH(data)), 0) FROM task_records) AS task_data_bytes,
         (SELECT COUNT(*) FROM task_history_records) AS history_rows,
         (SELECT COALESCE(SUM(LENGTH(data)), 0) FROM task_history_records) AS history_data_bytes,
         (SELECT COUNT(*) FROM r2_multipart_uploads) AS multipart_uploads,
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
  const lastBackup = await db
    .prepare("SELECT value, updated_at FROM app_settings WHERE key LIKE 'weekly_backup:%' ORDER BY updated_at DESC LIMIT 1")
    .first();
  let backup = null;
  try {
    backup = lastBackup?.value ? JSON.parse(lastBackup.value) : null;
  } catch {
    backup = lastBackup?.value ? { objectKey: lastBackup.value, generatedAt: Number(lastBackup.updated_at || 0) } : null;
  }
  const appDataBytes = Number(counts?.app_data_bytes || 0);
  const taskDataBytes = Number(counts?.task_data_bytes || 0);
  const historyDataBytes = Number(counts?.history_data_bytes || 0);
  const legacyBase64Bytes = Number(counts?.legacy_base64_bytes || 0);
  const temporaryBase64Bytes = Number(counts?.temporary_base64_bytes || 0);
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
        legacyBase64Bytes,
        temporaryBase64Bytes,
        appDataBytes,
        taskRows: Number(counts?.task_rows || 0),
        taskDataBytes,
        historyRows: Number(counts?.history_rows || 0),
        historyDataBytes,
        multipartUploads: Number(counts?.multipart_uploads || 0),
        approximatePayloadBytes:
          legacyBase64Bytes + temporaryBase64Bytes + appDataBytes + taskDataBytes + historyDataBytes,
        databaseBytes: pageCount && pageSize ? pageCount * pageSize : 0,
        reusableBytes: freePages && pageSize ? freePages * pageSize : 0,
        quotaBytes: 500_000_000,
        expiredSessions: Number(counts?.expired_sessions || 0),
        oldNotifications: Number(counts?.old_notifications || 0)
      },
      r2: {
        enabled: r2StorageEnabled(env),
        objects: r2Objects,
        bytes: r2Bytes,
        complete: r2Complete,
        freeTierBytes: 10_000_000_000
      },
      maintenance: {
        state: clean(maintenance?.value) || "pending",
        updatedAt: Number(maintenance?.updated_at || 0),
        failedMigrations: Number(failedMigrations?.count || 0)
      },
      backup
    }
  });
}

async function register(request, db) {
  const body = await readJson(request);
  const name = clean(body.name).toUpperCase();
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
  if (!isPrimaryCoordinatorUser(user)) {
    return json({ ok: false, message: "Solo Pablo puede cambiar el permiso para fechas anteriores." }, 403);
  }
  const body = await readJson(request);
  if (typeof body.enabled !== "boolean") {
    return json({ ok: false, message: "Selecciona si el permiso para sustentos y breaks anteriores debe estar encendido o apagado." }, 400);
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
      title: enabled ? "Fechas anteriores habilitadas" : "Fechas anteriores bloqueadas",
      body: enabled
        ? "Pablo habilito temporalmente la carga de sustentos y la modificacion de breaks de dias anteriores."
        : "Los sustentos y los breaks anteriores vuelven a estar protegidos.",
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

function userCanUploadTaskEvidence(user, task, allowPreviousDays = false) {
  return isSelfManagedMasterTask(user, task) || taskAllowsEvidenceUpload(task, dateIsoInLima(Date.now()), allowPreviousDays);
}

function validOfflineEvidenceQueue(task, value, now = Date.now()) {
  const queuedAt = Number(value || 0);
  return Boolean(
    Number.isFinite(queuedAt) &&
      queuedAt > 0 &&
      queuedAt <= now + 5 * 60 * 1000 &&
      now - queuedAt <= 7 * 24 * 60 * 60 * 1000 &&
      clean(task?.dueDate) === dateIsoInLima(queuedAt)
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
  const parentPath = `Task Hub / Sustentos / ${trainerName} / ${dueDate} / ${taskName}`;
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
      `SELECT chunk_index, LENGTH(chunk_base64) AS base64_length,
              SUBSTR(chunk_base64, -2) AS base64_tail
       FROM evidence_file_chunks WHERE file_id = ? ORDER BY chunk_index ASC`
    )
    .bind(file.id)
    .all();
  const chunkRows = chunks.results || [];
  if (!chunkRows.length || chunkRows.some((chunk, index) => Number(chunk.chunk_index) !== index)) {
    throw new Error("Incomplete legacy chunks");
  }
  const decodedLength = (chunk) => {
    const tail = String(chunk.base64_tail || "");
    const padding = tail.endsWith("==") ? 2 : tail.endsWith("=") ? 1 : 0;
    return Math.floor((Number(chunk.base64_length || 0) * 3) / 4) - padding;
  };
  const totalBytes = chunkRows.reduce((total, chunk) => total + decodedLength(chunk), 0);
  if (totalBytes < 1 || totalBytes > MAX_FILE_BYTES) throw new Error("Invalid legacy file size");

  const minimumPartBytes = 5 * 1024 * 1024;
  const groups = [];
  let currentGroup = [];
  let currentBytes = 0;
  for (const chunk of chunkRows) {
    currentGroup.push(chunk);
    currentBytes += decodedLength(chunk);
    if (currentBytes >= minimumPartBytes) {
      groups.push(currentGroup);
      currentGroup = [];
      currentBytes = 0;
    }
  }
  if (currentGroup.length) groups.push(currentGroup);

  const { objectKey, parentPath } = await r2EvidenceDescriptor(db, task, file);
  let multipartUpload;
  let completed = false;
  try {
    multipartUpload = await env.EVIDENCE_BUCKET.createMultipartUpload(objectKey, {
      httpMetadata: { contentType: file.mimeType },
      customMetadata: { fileId: clean(file.id), taskId: clean(task.id), ownerId: clean(task.ownerId) }
    });
    const uploadedParts = [];
    for (let groupIndex = 0; groupIndex < groups.length; groupIndex += 1) {
      const group = groups[groupIndex];
      const firstIndex = Number(group[0].chunk_index);
      const lastIndex = Number(group[group.length - 1].chunk_index);
      const rows = await db
        .prepare(
          `SELECT chunk_base64 FROM evidence_file_chunks
           WHERE file_id = ? AND chunk_index BETWEEN ? AND ? ORDER BY chunk_index ASC`
        )
        .bind(file.id, firstIndex, lastIndex)
        .all();
      if ((rows.results || []).length !== group.length) throw new Error("Legacy chunks changed during migration");
      const partBase64 = (rows.results || []).map((row) => String(row.chunk_base64 || "")).join("");
      const bytes = decodeEvidenceBase64(partBase64);
      if (groupIndex < groups.length - 1 && bytes.byteLength < minimumPartBytes) {
        throw new Error("Legacy multipart part is too small");
      }
      uploadedParts.push(await multipartUpload.uploadPart(groupIndex + 1, bytes));
    }
    const storedObject = await multipartUpload.complete(uploadedParts);
    completed = true;
    if (Number(storedObject?.size || 0) !== totalBytes) throw new Error("Incomplete R2 multipart migration");
    await db
      .prepare(
        `INSERT INTO evidence_storage
         (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
         VALUES (?, ?, ?, ?, ?, ?, '', ?)`
      )
      .bind(file.id, R2_STORAGE_PROVIDER, "EVIDENCE_BUCKET", objectKey, parentPath, totalBytes, Date.now())
      .run();
  } catch (error) {
    if (multipartUpload && !completed) await multipartUpload.abort().catch(() => {});
    if (completed) await env.EVIDENCE_BUCKET.delete(objectKey).catch(() => {});
    throw error;
  }
  return objectKey;
}

async function uploadEvidenceDirectToR2(request, db, user, env) {
  if (!r2StorageEnabled(env)) {
    return json({ ok: false, message: "El almacenamiento R2 no esta configurado." }, 503);
  }

  const url = new URL(request.url);
  const taskId = clean(url.searchParams.get("taskId"));
  const offlineQueuedAt = Number(url.searchParams.get("offlineQueuedAt") || 0);
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (clean(task.ownerId) !== clean(user.id) && !(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  const today = dateIsoInLima(Date.now());
  if (!userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled) && !validOfflineEvidenceQueue(task, offlineQueuedAt)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, today, data.lateEvidenceUploadsEnabled) }, 409);
  }

  const fileName = clean(url.searchParams.get("fileName")).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "archivo-sustento";
  const extension = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
  const mimeType = clean(request.headers.get("Content-Type")).split(";", 1)[0].toLowerCase() || mimeTypeForExtension(extension);
  const declaredSize = Number(url.searchParams.get("size") || request.headers.get("Content-Length") || 0);
  const contentLength = Number(request.headers.get("Content-Length") || url.searchParams.get("size") || 0);
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

async function initR2MultipartUpload(request, db, user, env) {
  if (!r2StorageEnabled(env)) {
    return json({ ok: false, message: "El almacenamiento R2 no esta configurado." }, 503);
  }
  const body = await readJson(request, 20_000);
  const taskId = clean(body.taskId);
  const offlineQueuedAt = Number(body.offlineQueuedAt || 0);
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (clean(task.ownerId) !== clean(user.id) && !(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  const today = dateIsoInLima(Date.now());
  if (!userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled) && !validOfflineEvidenceQueue(task, offlineQueuedAt)) {
    return json({ ok: false, message: evidenceUploadWindowError(task, today, data.lateEvidenceUploadsEnabled) }, 409);
  }

  const fileName = clean(body.fileName).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "archivo-sustento";
  const extension = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
  const mimeType = clean(body.mimeType).split(";", 1)[0].toLowerCase() || mimeTypeForExtension(extension);
  const sizeBytes = Number(body.size);
  const clientKey = clean(body.clientKey).slice(0, 220);
  if (
    !clientKey ||
    !ALLOWED_FILE_EXTENSIONS.has(extension) ||
    !ALLOWED_FILE_MIME_TYPES.has(mimeType) ||
    !Number.isInteger(sizeBytes) ||
    sizeBytes <= R2_MULTIPART_PART_BYTES ||
    sizeBytes > MAX_FILE_BYTES
  ) {
    return json({ ok: false, message: "El archivo no tiene un formato o tamano permitido." }, 400);
  }

  const existing = await db
    .prepare(
      `SELECT id, size_bytes, parts_json FROM r2_multipart_uploads
       WHERE submitted_by_id = ? AND task_id = ? AND client_key = ?`
    )
    .bind(user.id, taskId, clientKey)
    .first();
  if (existing && Number(existing.size_bytes) === sizeBytes) {
    return json({
      ok: true,
      uploadId: existing.id,
      partSize: R2_MULTIPART_PART_BYTES,
      uploadedParts: JSON.parse(existing.parts_json || "[]").map((part) => Number(part.partNumber))
    });
  }
  if (existing) {
    const stale = await db.prepare("SELECT object_key, r2_upload_id FROM r2_multipart_uploads WHERE id = ?").bind(existing.id).first();
    if (stale) await env.EVIDENCE_BUCKET.resumeMultipartUpload(stale.object_key, stale.r2_upload_id).abort().catch(() => {});
    await db.prepare("DELETE FROM r2_multipart_uploads WHERE id = ?").bind(existing.id).run();
  }

  const fileId = crypto.randomUUID();
  const descriptor = await r2EvidenceDescriptor(db, task, { id: fileId, fileName, mimeType });
  const multipart = await env.EVIDENCE_BUCKET.createMultipartUpload(descriptor.objectKey, {
    httpMetadata: { contentType: mimeType },
    customMetadata: { fileId, taskId: clean(task.id), ownerId: clean(task.ownerId) }
  });
  const now = Date.now();
  try {
    await db
      .prepare(
        `INSERT INTO r2_multipart_uploads
         (id, task_id, owner_id, submitted_by_id, client_key, file_name, mime_type, size_bytes,
          object_key, parent_path, r2_upload_id, parts_json, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '[]', ?, ?)`
      )
      .bind(
        fileId,
        task.id,
        task.ownerId,
        user.id,
        clientKey,
        fileName,
        mimeType,
        sizeBytes,
        descriptor.objectKey,
        descriptor.parentPath,
        multipart.uploadId,
        now,
        now
      )
      .run();
  } catch (error) {
    await multipart.abort().catch(() => {});
    throw error;
  }
  return json({ ok: true, uploadId: fileId, partSize: R2_MULTIPART_PART_BYTES, uploadedParts: [] }, 201);
}

async function uploadR2MultipartPart(request, db, user, env) {
  if (!r2StorageEnabled(env)) return json({ ok: false, message: "R2 no esta disponible." }, 503);
  const url = new URL(request.url);
  const uploadId = clean(url.searchParams.get("uploadId"));
  const partNumber = Number(url.searchParams.get("partNumber"));
  const row = await db.prepare("SELECT * FROM r2_multipart_uploads WHERE id = ?").bind(uploadId).first();
  if (!row) return json({ ok: false, message: "La carga vencio o ya fue completada." }, 404);
  if (clean(row.submitted_by_id) !== clean(user.id)) {
    return json({ ok: false, message: "No puedes continuar esta carga." }, 403);
  }
  const totalParts = Math.ceil(Number(row.size_bytes) / R2_MULTIPART_PART_BYTES);
  const expectedBytes = partNumber === totalParts
    ? Number(row.size_bytes) - R2_MULTIPART_PART_BYTES * (totalParts - 1)
    : R2_MULTIPART_PART_BYTES;
  const contentLength = Number(request.headers.get("Content-Length") || url.searchParams.get("size") || 0);
  if (!request.body || !Number.isInteger(partNumber) || partNumber < 1 || partNumber > totalParts || contentLength !== expectedBytes) {
    return json({ ok: false, message: "El bloque de carga esta incompleto." }, 400);
  }
  let uploadedPart;
  try {
    uploadedPart = await env.EVIDENCE_BUCKET
      .resumeMultipartUpload(row.object_key, row.r2_upload_id)
      .uploadPart(partNumber, request.body);
  } catch (error) {
    console.error("R2 multipart part", uploadId, partNumber, error);
    return json({ ok: false, message: "R2 no pudo recibir este bloque. Se reintentara sin empezar de cero." }, 503);
  }
  const parts = JSON.parse(row.parts_json || "[]").filter((part) => Number(part.partNumber) !== partNumber);
  parts.push({ partNumber, etag: uploadedPart.etag });
  parts.sort((left, right) => left.partNumber - right.partNumber);
  await db
    .prepare("UPDATE r2_multipart_uploads SET parts_json = ?, updated_at = ? WHERE id = ?")
    .bind(JSON.stringify(parts), Date.now(), uploadId)
    .run();
  return json({ ok: true, partNumber }, 201);
}

async function completeR2MultipartUpload(request, db, user, env) {
  if (!r2StorageEnabled(env)) return json({ ok: false, message: "R2 no esta disponible." }, 503);
  const body = await readJson(request, 10_000);
  const uploadId = clean(body.uploadId);
  const completedFile = await db
    .prepare("SELECT id, file_name, mime_type, created_at FROM evidence_files WHERE id = ?")
    .bind(uploadId)
    .first();
  if (completedFile) {
    return json({
      ok: true,
      file: {
        id: completedFile.id,
        name: completedFile.file_name,
        mimeType: completedFile.mime_type,
        createdAt: Number(completedFile.created_at),
        url: `/cloud/evidence/${completedFile.id}/file`
      }
    });
  }
  const row = await db.prepare("SELECT * FROM r2_multipart_uploads WHERE id = ?").bind(uploadId).first();
  if (!row) return json({ ok: false, message: "La carga vencio y debe iniciarse nuevamente." }, 404);
  if (clean(row.submitted_by_id) !== clean(user.id)) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const parts = JSON.parse(row.parts_json || "[]").sort((left, right) => left.partNumber - right.partNumber);
  const totalParts = Math.ceil(Number(row.size_bytes) / R2_MULTIPART_PART_BYTES);
  if (parts.length !== totalParts || parts.some((part, index) => Number(part.partNumber) !== index + 1 || !clean(part.etag))) {
    return json({ ok: false, message: "Aun faltan bloques del archivo por subir." }, 409);
  }

  let storedObject;
  try {
    storedObject = await env.EVIDENCE_BUCKET.resumeMultipartUpload(row.object_key, row.r2_upload_id).complete(parts);
  } catch (error) {
    storedObject = await env.EVIDENCE_BUCKET.head(row.object_key).catch(() => null);
    if (!storedObject || Number(storedObject.size || 0) !== Number(row.size_bytes)) {
      console.error("R2 multipart complete", uploadId, error);
      return json({ ok: false, message: "R2 no pudo finalizar el archivo. Vuelve a intentarlo." }, 503);
    }
  }
  if (Number(storedObject?.size || 0) !== Number(row.size_bytes)) {
    await env.EVIDENCE_BUCKET.delete(row.object_key).catch(() => {});
    return json({ ok: false, message: "La carga quedo incompleta y fue descartada." }, 409);
  }
  const createdAt = Date.now();
  try {
    await db.batch([
      db
        .prepare(
          `INSERT OR IGNORE INTO evidence_files
           (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at)
           VALUES (?, ?, ?, ?, ?, ?, '', ?)`
        )
        .bind(row.id, row.task_id, row.owner_id, row.submitted_by_id, row.file_name, row.mime_type, createdAt),
      db
        .prepare(
          `INSERT OR REPLACE INTO evidence_storage
           (file_id, provider, drive_id, drive_item_id, parent_path, size_bytes, sha256, created_at)
           VALUES (?, ?, ?, ?, ?, ?, '', ?)`
        )
        .bind(row.id, R2_STORAGE_PROVIDER, "EVIDENCE_BUCKET", row.object_key, row.parent_path, row.size_bytes, createdAt),
      db.prepare("DELETE FROM r2_multipart_uploads WHERE id = ?").bind(row.id)
    ]);
  } catch (error) {
    console.error("R2 multipart metadata", uploadId, error);
    return json({ ok: false, message: "El archivo llego a R2, pero falta registrar su confirmacion. Reintenta." }, 503);
  }
  return json({
    ok: true,
    file: { id: row.id, name: row.file_name, mimeType: row.mime_type, createdAt, url: `/cloud/evidence/${row.id}/file` }
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

async function r2SegmentedFileResponse(env, storage, row) {
  if (!r2StorageEnabled(env)) return json({ ok: false, message: "R2 no esta disponible." }, 503);
  const manifestObject = await env.EVIDENCE_BUCKET.get(clean(storage.drive_item_id));
  if (!manifestObject) return json({ ok: false, message: "Manifiesto de archivo no encontrado en R2." }, 404);
  let manifest;
  try {
    manifest = JSON.parse(await manifestObject.text());
  } catch {
    return json({ ok: false, message: "El manifiesto del archivo no es valido." }, 409);
  }
  const segments = Array.isArray(manifest.segments) ? manifest.segments : [];
  const totalSize = segments.reduce((total, segment) => total + Number(segment.size || 0), 0);
  if (!segments.length || totalSize !== Number(storage.size_bytes || manifest.size || 0)) {
    return json({ ok: false, message: "El archivo segmentado esta incompleto." }, 409);
  }
  let segmentIndex = 0;
  let activeReader = null;
  const stream = new ReadableStream({
    async pull(controller) {
      try {
        while (segmentIndex < segments.length) {
          if (!activeReader) {
            const object = await env.EVIDENCE_BUCKET.get(clean(segments[segmentIndex].key));
            if (!object) throw new Error("Missing R2 evidence segment");
            activeReader = object.body.getReader();
          }
          const chunk = await activeReader.read();
          if (chunk.done) {
            activeReader = null;
            segmentIndex += 1;
            continue;
          }
          controller.enqueue(chunk.value);
          return;
        }
        controller.close();
      } catch (error) {
        controller.error(error);
      }
    },
    async cancel() {
      if (activeReader) await activeReader.cancel().catch(() => {});
    }
  });
  const safeName = String(row.file_name || "sustento.jpg").replace(/["\r\n]/g, "-");
  const disposition =
    String(row.mime_type).startsWith("image/") ||
    String(row.mime_type).startsWith("video/") ||
    row.mime_type === "application/pdf"
      ? "inline"
      : "attachment";
  return new Response(stream, {
    headers: {
      "Content-Type": row.mime_type,
      "Content-Disposition": `${disposition}; filename="${safeName}"`,
      "Cache-Control": "private, max-age=300",
      "Content-Length": String(totalSize),
      "X-Content-Type-Options": "nosniff"
    }
  });
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
  const parentPath = `Task Hub / Sustentos / ${trainerName} / ${dueDate} / ${taskName}`;
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
  if (task.ownerId !== user.id && !(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  if (!userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled)) {
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
  if (task.ownerId !== user.id && !(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes subir sustentos para esa tarea." }, 403);
  }
  if (!userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled)) {
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
  if (row.submitted_by_id !== user.id && row.owner_id !== user.id && !(await coordinatorCanManageOwner(db, user, row.owner_id))) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === clean(row.task_id));
  if (!task || !userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled)) {
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
  if (row.submitted_by_id !== user.id && row.owner_id !== user.id && !(await coordinatorCanManageOwner(db, user, row.owner_id))) {
    return json({ ok: false, message: "No puedes completar esta carga." }, 403);
  }
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === clean(row.task_id));
  if (!task || !userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled)) {
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
  const coordinatorAllowed = await coordinatorCanManageOwner(db, user, row.owner_id);
  const allowed = coordinatorAllowed || observerAllowed || row.submitted_by_id === user.id || row.owner_id === user.id;
  if (!allowed) return json({ ok: false, message: "No tienes acceso a este archivo." }, 403);

  const storage = await db.prepare("SELECT * FROM evidence_storage WHERE file_id = ?").bind(row.id).first();
  if (storage?.provider === R2_STORAGE_PROVIDER) return r2FileResponse(env, storage, row);
  if (storage?.provider === R2_SEGMENTED_STORAGE_PROVIDER) return r2SegmentedFileResponse(env, storage, row);
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
    .prepare("SELECT id, name, email, role, status, access_level, team, job_title, member_type FROM users WHERE id = ? AND status = 'Activo'")
    .bind(ownerId)
    .first();
  if (!owner || isObserverUser(owner)) return json({ ok: false, message: "El responsable no esta disponible para tareas." }, 400);
  if (user.role === "Coordinador" && !coordinatorCanManageTaskUser(user, owner)) {
    return json({ ok: false, message: "Solo puedes asignar tareas dentro de tu equipo." }, 403);
  }
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
  const selfManagedMaster = isMasterUser(user) && clean(ownerId) === clean(user.id);
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
    status: selfManagedMaster ? "Cumplida" : "Pendiente",
    createdAt,
    completedAt: selfManagedMaster ? createdAt : 0,
    autoApproved: selfManagedMaster,
    history: [
      { type: "Asignacion", toId: ownerId, byId: user.id, reason: "Tarea creada", at: createdAt },
      ...(selfManagedMaster
        ? [{ type: "AutoaprobacionMaster", byId: user.id, reason: "Tarea propia registrada por un integrante Master", at: createdAt }]
        : [])
    ],
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

async function updateMasterTask(request, db, user) {
  const body = await readJson(request);
  const taskId = clean(body.taskId);
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (!isSelfManagedMasterTask(user, task)) {
    return json({ ok: false, message: "Solo el Master que registro la tarea puede editar su informacion." }, 403);
  }

  const title = clean(body.title).slice(0, 140);
  const category = normalizeTaskCategory(body.category);
  const product = clean(body.product).slice(0, 120) || "GENERAL";
  const description = clean(body.description).slice(0, 1500);
  const reason = clean(body.reason).slice(0, 500);
  if (title.length < 2 || reason.length < 3) {
    return json({ ok: false, message: "Completa el nombre y el motivo de la edicion." }, 400);
  }

  const previous = {
    title: clean(task.title),
    category: normalizeTaskCategory(task.category),
    product: clean(task.product) || "GENERAL",
    description: clean(task.description)
  };
  const next = { title, category, product, description };
  if (JSON.stringify(previous) === JSON.stringify(next)) {
    return json({ ok: false, message: "No hay cambios para guardar." }, 409);
  }

  Object.assign(task, next);
  task.history = Array.isArray(task.history) ? task.history : [];
  task.history.push({
    type: "EdicionMaster",
    byId: user.id,
    reason,
    previous,
    next,
    at: Date.now()
  });
  await saveData(db, data);
  return stateResponse(db, user, data);
}

async function submitTaskEvidence(request, db, user, context) {
  const body = await readJson(request, 1_000_000);
  const taskId = clean(body.taskId);
  const submittedEvidence = body.evidence || {};
  const offlineQueuedAt = Number(body.offlineQueuedAt || submittedEvidence.offlineQueuedAt || 0);
  const data = await loadData(db);
  const task = data.tasks.find((item) => clean(item.id) === taskId);
  if (!task) return json({ ok: false, message: "La tarea ya no existe." }, 404);
  if (clean(task.ownerId) !== user.id) {
    return json({ ok: false, message: "Solo el responsable de la tarea puede enviar el sustento." }, 403);
  }
  const selfManagedMaster = isSelfManagedMasterTask(user, task);
  if (task.status === "Cumplida" && !selfManagedMaster) {
    return json({ ok: false, message: "La tarea ya fue aprobada y completada." }, 409);
  }
  const validOfflineQueue = validOfflineEvidenceQueue(task, offlineQueuedAt);
  if (!userCanUploadTaskEvidence(user, task, data.lateEvidenceUploadsEnabled) && !validOfflineQueue) {
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
  const syncedAt = Date.now();
  const submittedAt = validOfflineQueue ? offlineQueuedAt : syncedAt;
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
    review: selfManagedMaster ? "Aprobada" : "Pendiente",
    reviewedAt: selfManagedMaster ? submittedAt : 0,
    reviewedById: selfManagedMaster ? user.id : "",
    reviewNote: selfManagedMaster ? "Autoaprobado por perfil Master." : ""
  };
  if (validOfflineQueue) {
    evidence.offlineQueuedAt = offlineQueuedAt;
    evidence.syncedAt = syncedAt;
  }
  if (lateAuthorization) evidence.lateAuthorization = lateAuthorization;
  task.evidence.push(evidence);
  task.history = Array.isArray(task.history) ? task.history : [];
  task.status = selfManagedMaster ? "Cumplida" : "En revision";
  if (selfManagedMaster) {
    task.completedAt = Number(task.completedAt || submittedAt);
    task.autoApproved = true;
  }
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
  if (validOfflineQueue) {
    task.history.push({
      type: "SustentoOffline",
      byId: user.id,
      reason: "Guardado en el dispositivo durante una interrupcion de Cloudflare y sincronizado posteriormente",
      at: syncedAt
    });
  }
  task.history.push({
    type: selfManagedMaster ? "SustentoMaster" : "Sustento",
    byId: user.id,
    fromId: user.id,
    reason: selfManagedMaster
      ? `${files.length} archivo${files.length === 1 ? "" : "s"} agregado${files.length === 1 ? "" : "s"} y autoaprobado${files.length === 1 ? "" : "s"}`
      : `${files.length} archivo${files.length === 1 ? "" : "s"} enviado${files.length === 1 ? "" : "s"}`,
    at: submittedAt
  });

  await saveData(db, data);
  const submitter = await db.prepare("SELECT team FROM users WHERE id = ?").bind(user.id).first();
  const coordinators = await db
    .prepare("SELECT id FROM users WHERE role = 'Coordinador' AND status = 'Activo' AND team = ?")
    .bind(userTeam(submitter || user))
    .all();
  const notifications = selfManagedMaster ? [] : (coordinators.results || [])
    .filter((coordinator) => clean(coordinator.id) !== user.id)
    .map((coordinator) => ({
      userId: coordinator.id,
      title: "Nuevo sustento por revisar",
      body: `${clean(task.title)} | ${clean(user.name)}`,
      url: `/?view=evidenceView&task=${encodeURIComponent(taskId)}`,
      sourceKey: `evidence:${taskId}:${evidenceId}:${clean(coordinator.id)}`
    }));
  try {
    const notifiedUsers = await queueNotifications(db, notifications);
    if (notifiedUsers.length) context.waitUntil(pushNotificationsForUsers(db, notifiedUsers));
  } catch (error) {
    console.error("Evidence notification queue", error);
  }
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
  if (!(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes autorizar tareas de otro equipo." }, 403);
  }
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
  if (!(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes revisar tareas de otro equipo." }, 403);
  }
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

async function deleteSegmentedR2EvidenceItems(env, items) {
  if (!r2StorageEnabled(env) || !items.length) return;
  for (const item of items) {
    const manifestKey = clean(item.drive_item_id);
    if (!manifestKey) continue;
    try {
      const manifestObject = await env.EVIDENCE_BUCKET.get(manifestKey);
      const manifest = manifestObject ? JSON.parse(await manifestObject.text()) : null;
      const keys = [manifestKey, ...(manifest?.segments || []).map((segment) => clean(segment.key)).filter(Boolean)];
      await env.EVIDENCE_BUCKET.delete(keys);
    } catch (error) {
      console.error("Segmented R2 evidence deletion", error);
    }
  }
}

async function deleteExternalEvidenceItems(env, items) {
  const r2Items = items.filter((item) => item.provider === R2_STORAGE_PROVIDER);
  const segmentedR2Items = items.filter((item) => item.provider === R2_SEGMENTED_STORAGE_PROVIDER);
  const oneDriveItems = items.filter((item) => item.provider === MICROSOFT_STORAGE_PROVIDER);
  await Promise.all([
    deleteR2EvidenceItems(env, r2Items),
    deleteSegmentedR2EvidenceItems(env, segmentedR2Items),
    deleteOneDriveEvidenceItems(env, oneDriveItems)
  ]);
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
  if (!(await coordinatorCanManageOwner(db, user, task.ownerId))) {
    return json({ ok: false, message: "No puedes eliminar tareas de otro equipo." }, 403);
  }
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
    .prepare("SELECT id, name, email, zone, role, status, created_at, access_level, team, job_title, member_type FROM users WHERE id = ? AND status = 'Activo'")
    .bind(targetUserId)
    .first();
  if (!target || target.role !== "Trainer" || isObserverUser(target)) return json({ ok: false, message: "Selecciona un trainer activo." }, 400);
  if (!coordinatorCanManageUser(user, target)) {
    return json({ ok: false, message: "Solo puedes ampliar la jornada de integrantes de tu equipo." }, 403);
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

  const coordinators = await db
    .prepare("SELECT id FROM users WHERE role = 'Coordinador' AND status = 'Activo' AND team = ?")
    .bind(userTeam(user))
    .all();
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
    .prepare("SELECT id, name, email, zone, role, status, created_at, access_level, team, job_title, member_type FROM users WHERE id = ? AND status = 'Activo'")
    .bind(overtimeRequest.userId)
    .first();
  if (!target || target.role !== "Trainer" || isObserverUser(target)) return json({ ok: false, message: "El trainer ya no esta activo." }, 400);
  if (!coordinatorCanManageUser(user, target)) {
    return json({ ok: false, message: "No puedes revisar solicitudes de otro equipo." }, 403);
  }
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
  const userRows = await db
    .prepare(
      `SELECT id, name, email, zone, role, status, created_at, access_level, team, job_title, member_type
       FROM users ORDER BY created_at ASC`
    )
    .all();
  const allUserRows = userRows.results || [];
  const visibleUserRows = observer
    ? allUserRows
    : coordinator
      ? allUserRows.filter((row) => userTeam(row) === userTeam(user))
      : allUserRows.filter((row) => row.status === "Activo" && userTeam(row) === userTeam(user));
  const visibleOwnerIds = new Set(visibleUserRows.map((row) => clean(row.id)));
  const visibleTaskOwnerIds = coordinator
    ? new Set(visibleUserRows.filter((row) => coordinatorCanManageTaskUser(user, row)).map((row) => clean(row.id)))
    : visibleOwnerIds;
  const users = visibleUserRows.map(publicUser);
  const visibleTasks = observer
    ? data.tasks
    : coordinator
      ? data.tasks.filter((task) => visibleTaskOwnerIds.has(clean(task.ownerId)))
      : data.tasks.filter((task) => clean(task.ownerId) === clean(user.id));
  const visibleDeletedTasks = observer
    ? data.deletedTasks
    : coordinator
      ? (data.deletedTasks || []).filter((task) => visibleTaskOwnerIds.has(clean(task.ownerId)))
      : (data.deletedTasks || []).filter((task) => clean(task.ownerId) === clean(user.id));
  const visibleSchedules = observer
    ? data.workScheduleByUserDate
    : Object.fromEntries(
        Object.entries(data.workScheduleByUserDate || {}).filter(([userId]) => visibleOwnerIds.has(clean(userId)))
      );
  const visibleOvertime = observer
    ? data.overtimeRequests
    : coordinator
      ? (data.overtimeRequests || []).filter((item) => visibleOwnerIds.has(clean(item.userId)))
      : (data.overtimeRequests || []).filter((item) => clean(item.userId) === clean(user.id));
  const state = {
    ...EMPTY_DATA,
    ...data,
    activeUserId: user.id,
    users,
    tasks: visibleTasks,
    deletedTasks: visibleDeletedTasks,
    workScheduleByUserDate: visibleSchedules,
    overtimeRequests: visibleOvertime,
    announcements: coordinator || observer
      ? data.announcements
      : data.announcements.filter((item) => item.audience === "all" || item.targetId === user.id),
    supportRequests: observer
      ? data.supportRequests
      : (data.supportRequests || []).filter(
          (item) =>
            clean(item.fromId) === clean(user.id) ||
            clean(item.toId) === clean(user.id) ||
            (coordinator && (visibleOwnerIds.has(clean(item.fromId)) || visibleOwnerIds.has(clean(item.toId))))
        ),
    passwordRecoveryRequests: coordinator
      ? (data.passwordRecoveryRequests || []).filter((item) => {
          const target = allUserRows.find((row) => clean(row.email).toLowerCase() === clean(item.email).toLowerCase());
          return !target || userTeam(target) === userTeam(user);
        })
      : [],
    registrationRequests: coordinator
      ? (data.registrationRequests || []).filter((item) => clean(item.team || TEAM_TRAINING) === userTeam(user))
      : []
  };
  return json({ ok: true, state, user: publicUser(user) });
}

async function putState(request, db, user, context, env) {
  const body = await readJson(request, 6_000_000);
  const submitted = body.state || {};
  const current = await loadData(db);
  const notifications = [];
  const activeUserRows = await db
    .prepare("SELECT id, name, email, role, access_level, team, job_title, member_type FROM users WHERE status = 'Activo'")
    .all();
  const activeUsers = activeUserRows.results || [];
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
    const today = dateIsoInLima(Date.now());
    const conflictFreeDates = Object.fromEntries(
      Object.entries(normalizedDates[user.id] || {}).filter(
        ([dateValue, settings]) =>
          (dateValue >= today || current.lateEvidenceUploadsEnabled) &&
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
    const managedUserIds = new Set(
      activeUsers.filter((row) => coordinatorCanManageTaskUser(user, row)).map((row) => clean(row.id))
    );
    if (isPrimaryCoordinatorUser(user)) {
      for (const key of ["announcements", "dailyMotivations"]) {
        if (Array.isArray(submitted[key])) current[key] = submitted[key];
      }
    }
    if (Array.isArray(submitted.registrationRequests)) {
      current.registrationRequests = mergeScopedItems(
        current.registrationRequests,
        submitted.registrationRequests,
        (item) => clean(item.team || TEAM_TRAINING) === userTeam(user)
      );
    }
    if (Array.isArray(submitted.passwordRecoveryRequests)) {
      current.passwordRecoveryRequests = mergeScopedItems(
        current.passwordRecoveryRequests,
        submitted.passwordRecoveryRequests,
        (item) => managedUserIds.has(clean(item.userId))
      );
    }
    if (Array.isArray(submitted.supportRequests)) {
      current.supportRequests = mergeScopedItems(
        current.supportRequests,
        submitted.supportRequests,
        (item) => managedUserIds.has(clean(item.fromId)) || managedUserIds.has(clean(item.toId))
      );
    }
    if (Array.isArray(submitted.tasks)) {
      const previousManagedTasks = current.tasks.filter((task) => managedUserIds.has(clean(task.ownerId)));
      const protectedTasks = current.tasks.filter((task) => !managedUserIds.has(clean(task.ownerId)));
      const submittedManagedTasks = submitted.tasks.filter((task) => managedUserIds.has(clean(task.ownerId)));
      const previousTasks = new Map(previousManagedTasks.map((task) => [clean(task.id), task]));
      const submittedTaskIds = new Set(submittedManagedTasks.map((task) => clean(task.id)).filter(Boolean));
      const removedTasks = previousManagedTasks.filter((task) => clean(task.id) && !submittedTaskIds.has(clean(task.id)));
      const updatedManagedTasks = submittedManagedTasks
        .map((task) => {
          const normalizedTask = {
            ...task,
            category: normalizeTaskCategory(task.category)
          };
          const previous = previousTasks.get(clean(normalizedTask.id));
          if (!previous && (!managedUserIds.has(clean(normalizedTask.ownerId)) || clean(normalizedTask.title).length < 2)) return null;
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
      current.tasks = [...protectedTasks, ...updatedManagedTasks];
      for (const task of updatedManagedTasks) {
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
        await deleteTaskEvidenceFiles(db, clean(removedTask.id), env, context);
      }
    }
  } else {
    const submittedTasks = new Map((submitted.tasks || []).map((task) => [task.id, task]));
    const activeTrainerIds = new Set(
      activeUsers
        .filter(
          (row) =>
            row.role === "Trainer" &&
            !isObserverUser(row) &&
            userMemberType(row) !== "master" &&
            userTeam(row) === userTeam(user)
        )
        .map((row) => row.id)
    );
    current.tasks = current.tasks.map((task) => {
      if (task.ownerId !== user.id) return task;
      const next = submittedTasks.get(task.id);
      if (!next) return task;
      const requestedOwnerId = clean(next.ownerId);
      const ownerChanged =
        !isMasterUser(user) &&
        requestedOwnerId &&
        requestedOwnerId !== task.ownerId &&
        activeTrainerIds.has(requestedOwnerId);
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

function mergeScopedItems(currentItems, submittedItems, isInScope) {
  const protectedItems = (currentItems || []).filter((item) => !isInScope(item));
  const scopedItems = (submittedItems || []).filter(isInScope);
  return [...protectedItems, ...scopedItems];
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
        clean(item.title).slice(0, 100) || "Task Hub",
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
        console.error("Task Hub push", error);
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
  const name = clean(body.name).toUpperCase();
  const team = userTeam(actor);
  const jobTitle = clean(body.jobTitle).slice(0, 120) || (team === TEAM_AUDIOVISUAL ? "Creador de Contenido" : "Trainer");
  const memberType = team === TEAM_AUDIOVISUAL ? "audiovisual" : "trainer";
  await db
    .prepare(
      `INSERT INTO users
       (id, name, email, zone, role, status, password_hash, password_salt, created_at, access_level, team, job_title, member_type)
       VALUES (?, ?, ?, ?, 'Trainer', 'Activo', ?, ?, ?, '', ?, ?, ?)`
    )
    .bind(userId, name, email, clean(body.zone), passwordData.hash, passwordData.salt, createdAt, team, jobTitle, memberType)
    .run();
  return json({
    ok: true,
    user: publicUser({
      id: userId,
      name,
      email,
      zone: clean(body.zone),
      role: "Trainer",
      status: "Activo",
      team,
      job_title: jobTitle,
      member_type: memberType,
      createdAt
    })
  }, 201);
}

async function resetPassword(request, db, actor) {
  if (actor.role !== "Coordinador") return json({ ok: false, message: "Permiso insuficiente." }, 403);
  const body = await readJson(request);
  const user = await db
    .prepare("SELECT id, name, email, role, status, access_level, team, member_type FROM users WHERE id = ?")
    .bind(clean(body.userId))
    .first();
  if (!user) return json({ ok: false, message: "Usuario no encontrado." }, 404);
  if (!coordinatorCanManageUser(actor, user)) {
    return json({ ok: false, message: "No puedes restablecer accesos de otro equipo." }, 403);
  }
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
  const [appResult, taskResult, historyResult] = await db.batch([
    db.prepare("SELECT data FROM app_data WHERE id = 1"),
    db.prepare("SELECT id, archived, data FROM task_records ORDER BY archived ASC, created_at DESC, id ASC"),
    db.prepare("SELECT task_id, entry_index, data FROM task_history_records ORDER BY task_id ASC, entry_index ASC")
  ]);
  try {
    const parsed = JSON.parse(appResult?.results?.[0]?.data || "{}");
    const historyByTask = new Map();
    for (const row of historyResult?.results || []) {
      try {
        const history = historyByTask.get(row.task_id) || [];
        history[Number(row.entry_index)] = JSON.parse(row.data || "{}");
        historyByTask.set(row.task_id, history);
      } catch {
        console.error("Invalid task history row", row.task_id, row.entry_index);
      }
    }
    const storedTasks = [];
    const storedDeletedTasks = [];
    for (const row of taskResult?.results || []) {
      try {
        const task = {
          ...JSON.parse(row.data || "{}"),
          id: row.id,
          history: (historyByTask.get(row.id) || []).filter(Boolean)
        };
        (Number(row.archived) ? storedDeletedTasks : storedTasks).push(task);
      } catch {
        console.error("Invalid task record", row.id);
      }
    }
    const tasks = storedTasks.length || storedDeletedTasks.length ? storedTasks : parsed.tasks || [];
    const deletedTasks = storedTasks.length || storedDeletedTasks.length ? storedDeletedTasks : parsed.deletedTasks || [];
    const data = {
      ...structuredClone(EMPTY_DATA),
      ...parsed,
      version: 23,
      workSettings: normalizeWorkSettings(parsed.workSettings),
      breakSettingsByUser: normalizeBreakSettingsByUser(parsed.breakSettingsByUser),
      breakSettingsByUserDate: normalizeBreakSettingsByUserDate(parsed.breakSettingsByUserDate),
      workScheduleByUserDate: normalizeWorkScheduleByUserDate(parsed.workScheduleByUserDate),
      overtimeRequests: normalizeOvertimeRequests(parsed.overtimeRequests),
      tasks: tasks.map((task) => ({ ...task, category: normalizeTaskCategory(task.category) })),
      deletedTasks: deletedTasks.map((task) => ({ ...task, category: normalizeTaskCategory(task.category) })),
      lgUpdates: normalizeLgUpdates(parsed.lgUpdates),
      lateEvidenceUploadsEnabled: Boolean(parsed.lateEvidenceUploadsEnabled),
      lateEvidencePolicyHistory: Array.isArray(parsed.lateEvidencePolicyHistory)
        ? parsed.lateEvidencePolicyHistory.slice(0, 100)
        : []
    };
    Object.defineProperty(data, "__taskSnapshot", {
      configurable: true,
      enumerable: false,
      value: createTaskSnapshot(data)
    });
    return data;
  } catch {
    const data = structuredClone(EMPTY_DATA);
    Object.defineProperty(data, "__taskSnapshot", {
      configurable: true,
      enumerable: false,
      value: new Map()
    });
    return data;
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
    announcements: data.announcements || [],
    supportRequests: data.supportRequests || [],
    dailyMotivations: data.dailyMotivations || [],
    lgUpdates: normalizeLgUpdates(data.lgUpdates),
    lateEvidenceUploadsEnabled: Boolean(data.lateEvidenceUploadsEnabled),
    lateEvidencePolicyHistory: Array.isArray(data.lateEvidencePolicyHistory)
      ? data.lateEvidencePolicyHistory.slice(0, 100)
      : []
  };
  const now = Date.now();
  const previous = data.__taskSnapshot instanceof Map ? data.__taskSnapshot : new Map();
  const current = createTaskSnapshot(data);
  const statements = [
    db.prepare("UPDATE app_data SET data = ?, updated_at = ? WHERE id = 1").bind(JSON.stringify(payload), now)
  ];

  for (const [taskId, snapshot] of current) {
    const before = previous.get(taskId);
    if (!before || before.record !== snapshot.record || before.archived !== snapshot.archived) {
      statements.push(
        db
          .prepare(
            `INSERT INTO task_records (id, owner_id, due_date, status, archived, created_at, data, updated_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)
             ON CONFLICT(id) DO UPDATE SET
               owner_id = excluded.owner_id,
               due_date = excluded.due_date,
               status = excluded.status,
               archived = excluded.archived,
               created_at = excluded.created_at,
               data = excluded.data,
               updated_at = excluded.updated_at`
          )
          .bind(
            taskId,
            clean(snapshot.task.ownerId),
            clean(snapshot.task.dueDate),
            clean(snapshot.task.status),
            snapshot.archived,
            Number(snapshot.task.deletedAt || snapshot.task.createdAt || now),
            snapshot.record,
            now
          )
      );
    }
    if (!before || before.history !== snapshot.history) {
      statements.push(db.prepare("DELETE FROM task_history_records WHERE task_id = ?").bind(taskId));
      if (snapshot.history !== "[]") {
        statements.push(
          db
            .prepare(
              `INSERT INTO task_history_records (task_id, entry_index, data)
               SELECT ?, CAST(key AS INTEGER), value FROM json_each(?)`
            )
            .bind(taskId, snapshot.history)
        );
      }
    }
  }

  for (const taskId of previous.keys()) {
    if (current.has(taskId)) continue;
    statements.push(db.prepare("DELETE FROM task_history_records WHERE task_id = ?").bind(taskId));
    statements.push(db.prepare("DELETE FROM task_records WHERE id = ?").bind(taskId));
  }

  await db.batch(statements);
  Object.defineProperty(data, "__taskSnapshot", {
    configurable: true,
    enumerable: false,
    value: current
  });
}

function createTaskSnapshot(data) {
  const snapshot = new Map();
  const addTasks = (tasks, archived) => {
    for (const source of Array.isArray(tasks) ? tasks : []) {
      const taskId = clean(source?.id);
      if (!taskId) continue;
      const task = { ...source };
      const history = Array.isArray(task.history) ? task.history : [];
      delete task.history;
      snapshot.set(taskId, {
        task: source,
        archived,
        record: JSON.stringify(task),
        history: JSON.stringify(history)
      });
    }
  };
  addTasks(data.tasks, 0);
  addTasks(data.deletedTasks, 1);
  return snapshot;
}

function isObserverUser(user) {
  return clean(user?.access_level).toLowerCase() === OBSERVER_ACCESS_LEVEL;
}

function userTeam(user) {
  if (isObserverUser(user)) return "Todos";
  return clean(user?.team) || TEAM_TRAINING;
}

function userMemberType(user) {
  const value = clean(user?.member_type || user?.memberType).toLowerCase();
  if (value) return value;
  return user?.role === "Coordinador" ? "master" : "trainer";
}

function isMasterUser(user) {
  return userMemberType(user) === "master";
}

function isSelfManagedMasterTask(user, task) {
  return Boolean(
    isMasterUser(user) &&
      clean(task?.ownerId) === clean(user?.id) &&
      clean(task?.createdById) === clean(user?.id)
  );
}

function isPrimaryCoordinatorUser(user) {
  return user?.role === "Coordinador" && clean(user?.email).toLowerCase() === PRIMARY_COORDINATOR_EMAIL;
}

function coordinatorCanManageUser(actor, target) {
  return (
    actor?.role === "Coordinador" &&
    !isObserverUser(actor) &&
    target &&
    !isObserverUser(target) &&
    userTeam(actor) === userTeam(target)
  );
}

function coordinatorCanManageTaskUser(actor, target) {
  return Boolean(
    coordinatorCanManageUser(actor, target) &&
      (userMemberType(target) !== "master" || clean(actor?.id) === clean(target?.id))
  );
}

async function coordinatorCanManageOwner(db, actor, ownerId) {
  if (actor?.role !== "Coordinador" || !clean(ownerId)) return false;
  const target = await db
    .prepare("SELECT id, email, role, status, access_level, team, member_type FROM users WHERE id = ? AND status = 'Activo'")
    .bind(clean(ownerId))
    .first();
  return coordinatorCanManageTaskUser(actor, target);
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
    team: userTeam(row),
    jobTitle: clean(row.job_title || row.jobTitle) || (observer ? "Admin" : row.role === "Coordinador" ? "Coordinador" : "Trainer"),
    memberType: observer ? "admin" : userMemberType(row),
    status: row.status || "Activo",
    createdAt: Number(row.createdAt || row.created_at || Date.now())
  };
}

async function readJson(request, maxLength = 100000) {
  const text = await request.text();
  if (text.length > maxLength) {
    const error = new Error("La solicitud es demasiado grande.");
    error.status = 413;
    throw error;
  }
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    const error = new Error("El formato de la solicitud no es valido.");
    error.status = 400;
    throw error;
  }
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

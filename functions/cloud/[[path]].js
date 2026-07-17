const SESSION_COOKIE = "lg_session";
const SESSION_DAYS = 14;
const COORDINATOR_CODE_HASH = "e62163b1947feab8e4db70a99cffd5fb9c9f66d5e8901a4fb9775180ea780b71";

const EMPTY_DATA = {
  version: 5,
  registrationRequests: [],
  passwordRecoveryRequests: [],
  tasks: [],
  announcements: [],
  supportRequests: [],
  dailyMotivations: []
};

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

    if (route === "evidence/upload" && request.method === "POST") return uploadEvidence(request, env.DB, session.user);
    const evidencePhotoMatch = route.match(/^evidence\/([^/]+)\/photo$/);
    if (evidencePhotoMatch && request.method === "GET") return evidencePhoto(env.DB, session.user, evidencePhotoMatch[1]);
    if (route === "state" && request.method === "GET") return getState(env.DB, session.user);
    if (route === "state" && request.method === "PUT") return putState(request, env.DB, session.user);
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
    db.prepare("INSERT OR IGNORE INTO app_data (id, data, updated_at) VALUES (1, ?, 0)").bind(JSON.stringify(EMPTY_DATA)),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_sessions_expiry ON sessions(expires_at)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_task ON evidence_files(task_id)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_evidence_submitter ON evidence_files(submitted_by_id)")
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

async function getState(db, user) {
  return stateResponse(db, user);
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

  const photoMatch = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/.exec(String(body.photoData || ""));
  if (!photoMatch) return json({ ok: false, message: "La foto no tiene un formato valido." }, 400);
  if (photoMatch[2].length > 1_800_000) {
    return json({ ok: false, message: "La foto es demasiado grande. Intenta tomarla nuevamente." }, 413);
  }

  const fileId = crypto.randomUUID();
  const fileName = clean(body.fileName).replace(/[\\/:*?"<>|]/g, "-").slice(0, 120) || "sustento.jpg";
  const createdAt = Date.now();
  await db
    .prepare(
      "INSERT INTO evidence_files (id, task_id, owner_id, submitted_by_id, file_name, mime_type, photo_base64, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    )
    .bind(fileId, task.id, task.ownerId, user.id, fileName, photoMatch[1], photoMatch[2], createdAt)
    .run();

  return json({
    ok: true,
    file: { id: fileId, name: fileName, createdAt, url: `/cloud/evidence/${fileId}/photo` }
  }, 201);
}

async function evidencePhoto(db, user, fileId) {
  const row = await db.prepare("SELECT * FROM evidence_files WHERE id = ?").bind(clean(fileId)).first();
  if (!row) return json({ ok: false, message: "Archivo no encontrado." }, 404);
  const allowed = user.role === "Coordinador" || row.submitted_by_id === user.id || row.owner_id === user.id;
  if (!allowed) return json({ ok: false, message: "No tienes acceso a este archivo." }, 403);

  const binary = atob(row.photo_base64);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const safeName = String(row.file_name || "sustento.jpg").replace(/["\r\n]/g, "-");
  return new Response(bytes, {
    headers: {
      "Content-Type": row.mime_type,
      "Content-Disposition": `inline; filename="${safeName}"`,
      "Cache-Control": "private, max-age=300",
      "X-Content-Type-Options": "nosniff"
    }
  });
}

async function stateResponse(db, user) {
  const data = await loadData(db);
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

async function putState(request, db, user) {
  const body = await readJson(request, 6_000_000);
  const submitted = body.state || {};
  const current = await loadData(db);

  if (user.role === "Coordinador") {
    for (const key of ["registrationRequests", "passwordRecoveryRequests", "tasks", "announcements", "supportRequests", "dailyMotivations"]) {
      if (Array.isArray(submitted[key])) current[key] = submitted[key];
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
        clean(lastHistory.reason);
      return {
        ...task,
        ownerId: validReassignment ? requestedOwnerId : task.ownerId,
        status: validReassignment ? "Pendiente" : clean(next.status) || task.status,
        evidence: Array.isArray(next.evidence) ? next.evidence : task.evidence,
        history: validReassignment ? nextHistory : Array.isArray(next.history) ? next.history : task.history,
        blockedReason: validReassignment ? "" : clean(next.blockedReason),
        blockedAt: validReassignment ? 0 : Number(next.blockedAt || 0)
      };
    });
    const knownSupport = new Set(current.supportRequests.map((item) => item.id));
    for (const item of submitted.supportRequests || []) {
      if (!knownSupport.has(item.id) && item.fromId === user.id) current.supportRequests.unshift(item);
    }
  }

  await saveData(db, current);
  return stateResponse(db, user);
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

async function loadData(db) {
  const row = await db.prepare("SELECT data FROM app_data WHERE id = 1").first();
  try {
    return { ...structuredClone(EMPTY_DATA), ...JSON.parse(row?.data || "{}") };
  } catch {
    return structuredClone(EMPTY_DATA);
  }
}

async function saveData(db, data) {
  const payload = {
    version: 5,
    registrationRequests: data.registrationRequests || [],
    passwordRecoveryRequests: data.passwordRecoveryRequests || [],
    tasks: data.tasks || [],
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

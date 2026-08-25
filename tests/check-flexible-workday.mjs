import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");
const cloud = fs.readFileSync(new URL("../functions/cloud/[[path]].js", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html deben permanecer identicos");
if (cloud !== api) throw new Error("Las rutas cloud y api deben mantener la misma validacion");

const frontendMarkers = [
  'const EARLIEST_OVERTIME_START = "05:00"',
  'id="overtimeStart"',
  "function overtimeStartOptions",
  "startTime: clean(settings.startTime)",
  "start: extendedStart",
  "body: JSON.stringify({ userId, date: dateValue, startTime, endTime, reason })",
  "timeToMinutes(coordinatorScheduleBounds(dateValue).start)"
];

const backendMarkers = [
  "function baseWorkdayStart",
  "function workScheduleStartForUser",
  "workStartOverride = \"\"",
  "startTime: requestedStart",
  "timeToMinutes(startTime) < timeToMinutes(EARLIEST_OVERTIME_START)",
  "jornada ${startTime}-${endTime}"
];

for (const marker of frontendMarkers) {
  if (!html.includes(marker)) throw new Error(`Falta jornada flexible en frontend: ${marker}`);
}
for (const marker of backendMarkers) {
  if (!cloud.includes(marker)) throw new Error(`Falta jornada flexible en backend: ${marker}`);
}

const helperSource = cloud.slice(
  cloud.indexOf("function baseWorkdayStart"),
  cloud.indexOf("function hasTaskConflict")
);
const helpers = new Function(
  "WORKDAY_START",
  "WEEKDAY_END",
  "SATURDAY_END",
  "DANNY_SATURDAY_START",
  "DANNY_SATURDAY_END",
  "EARLIEST_OVERTIME_START",
  "normalizedUserName",
  "timeToMinutes",
  "clean",
  "normalizeWorkSettings",
  `${helperSource}; return { baseWorkdayStart, baseWorkdayEnd, workScheduleStartForUser, workScheduleEndForUser, validWorkSchedule };`
)(
  "08:30",
  "19:00",
  "14:00",
  "09:00",
  "14:30",
  "05:00",
  (value) => String(value || "").trim().toUpperCase(),
  (value) => {
    const [hours, minutes] = String(value || "").split(":").map(Number);
    return hours * 60 + minutes;
  },
  (value) => String(value || "").trim(),
  (settings) => settings || { breakStart: "12:30", breakEnd: "14:00" }
);

const monday = "2026-08-24";
const saturday = "2026-08-29";
const data = {
  workScheduleByUserDate: {
    trainer: {
      [monday]: { startTime: "07:00", endTime: "20:00" },
      [saturday]: { startTime: "07:00", endTime: "14:30" }
    }
  }
};
const noConflictBreak = { breakStart: "12:30", breakEnd: "14:00" };

if (helpers.workScheduleStartForUser(data, "trainer", monday, "TRAINER") !== "07:00") {
  throw new Error("La entrada anticipada 07:00 no fue aplicada");
}
if (helpers.workScheduleEndForUser(data, "trainer", monday, "TRAINER") !== "20:00") {
  throw new Error("La salida extendida 20:00 no fue aplicada");
}
if (!helpers.validWorkSchedule(monday, "07:00", "08:00", noConflictBreak, "TRAINER", "20:00", "07:00")) {
  throw new Error("El servidor rechazaria una tarea autorizada a las 07:00");
}
if (helpers.validWorkSchedule(monday, "06:45", "07:15", noConflictBreak, "TRAINER", "20:00", "07:00")) {
  throw new Error("El servidor permitiria una tarea anterior a la hora autorizada");
}
if (!helpers.validWorkSchedule(saturday, "07:00", "08:00", noConflictBreak, "DANNY DIOS", "14:30", "07:00")) {
  throw new Error("La entrada anticipada de Danny el sabado no fue respetada");
}

console.log("JornadaFlexible=05:00-23:45");
console.log("IngresoEjemplo=07:00");

import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");

const markers = [
  "--task-assigned: #2563eb",
  "--task-review: #9a6500",
  "--task-approved: #1f7a52",
  "--task-rejected: #c6283d",
  ".pill.pending,",
  ".pill.process {",
  ".pill.review,",
  ".pill.medium {",
  "background: var(--task-assigned-soft)",
  "background: var(--task-review-soft)",
  "background: var(--task-approved-soft)",
  "background: var(--task-rejected-soft)",
  "card.className = `trainer-task-card ${taskStatusClass(task.status)}`",
  "card.className = `item task-item ${taskStatusClass(task.status)}",
  'if (status === "En revision") return { fill: "FFFFF3D6"',
  'if (status === "Pendiente" || status === "En proceso") return { fill: "FFEAF2FF"'
];

for (const marker of markers) {
  if (!html.includes(marker)) throw new Error(`Falta el color de estado: ${marker}`);
}

console.log("Assigned=Blue");
console.log("Review=LightMustard");
console.log("Approved=Green");
console.log("Rejected=Red");

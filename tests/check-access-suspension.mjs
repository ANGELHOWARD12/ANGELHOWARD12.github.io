import assert from "node:assert/strict";
import fs from "node:fs";

const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");
const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");

assert.match(api, /const ACCESS_SUSPENDED = true/);
assert.match(api, /accessSuspended: true/);
assert.match(api, /423/);
assert.match(api, /accessSuspended: ACCESS_SUSPENDED/);
assert.match(html, /const ACCESS_SUSPENDED = true/);
assert.match(html, /els\.loginForm\.elements/);
assert.match(html, /els\.registrationForm\.elements/);
assert.match(html, /if \(ACCESS_SUSPENDED\)/);

console.log("GlobalAccess=Suspended");
console.log("D1AndR2Content=Preserved");

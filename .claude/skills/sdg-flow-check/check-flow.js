// Reads a Power Automate export (.zip) and prints every action plus known traps.
// Usage: node check-flow.js "PowerAutomateFlowsSDG/<name>.zip"
const { execSync } = require("child_process");

const zipPath = process.argv[2];
if (!zipPath) {
  console.log('Usage: node check-flow.js "PowerAutomateFlowsSDG/<name>.zip"');
  process.exit(1);
}

const listing = execSync(`unzip -Z1 "${zipPath}"`).toString().split("\n").map((line) => line.trim());
const definitionFile = listing.find((line) => line.endsWith("/definition.json"));
const raw = execSync(`unzip -p "${zipPath}" "${definitionFile}"`, { maxBuffer: 50 * 1024 * 1024 }).toString();
const flow = JSON.parse(raw).properties.definition;
const warnings = [];
const libraries = new Set();
const accounts = new Set();

function checkText(actionName, text) {
  if (typeof text !== "string") return;
  if (text.startsWith("@") && /\s$/.test(text)) {
    warnings.push(`${actionName}: expression ends with a space or line break`);
  }
  if (/\b(eq|ne)\s+(true|false)\b/.test(text)) {
    warnings.push(`${actionName}: filter compares a Yes/No column to true/false (use 1/0)`);
  }
  const titles = text.match(/getbytitle\('([^']+)'\)/g) || [];
  titles.forEach((t) => libraries.add(t.slice(12, -2)));
  const emails = text.match(/[a-z0-9.]+@sdguthrie\.com/gi) || [];
  emails.forEach((e) => accounts.add(e.toLowerCase()));
}

function scanValues(actionName, value) {
  if (typeof value === "string") return checkText(actionName, value);
  if (value && typeof value === "object") {
    Object.keys(value).forEach((key) => {
      checkText(actionName, key);
      scanValues(actionName, value[key]);
    });
  }
}

function checkAction(name, action) {
  const inputs = action.inputs || {};
  const params = inputs.parameters || {};

  // HTTP requests to SharePoint should ask for the nometadata shape.
  if (params.uri || params["parameters/uri"]) {
    const headers = JSON.stringify(params.headers || params["parameters/headers"] || "");
    if (!headers.includes("nometadata")) {
      warnings.push(`${name}: HTTP request without "Accept: application/json;odata=nometadata"`);
    }
  }

  // A Select must map to a plain value, not a key/value object.
  if (action.type === "Select" && typeof inputs.select === "object") {
    warnings.push(`${name}: Select is in key/value mode (expected a plain value)`);
  }

  // A switch case value must be a typed literal, never an expression.
  if (action.type === "Switch") {
    Object.keys(action.cases || {}).forEach((caseName) => {
      const caseValue = action.cases[caseName].case;
      if (typeof caseValue === "string" && caseValue.startsWith("@")) {
        warnings.push(`${name}: case "${caseName}" uses an expression instead of a literal`);
      }
      if (caseValue === "" || caseValue === undefined) {
        warnings.push(`${name}: case "${caseName}" has an empty value`);
      }
    });
  }

  scanValues(name, inputs);
  if (action.expression) scanValues(name, action.expression);
}

function walk(actions, depth) {
  Object.keys(actions || {}).forEach((name) => {
    const action = actions[name];
    const runAfter = Object.keys(action.runAfter || {})
      .map((k) => `${k}[${action.runAfter[k].join(",")}]`)
      .join(" ");
    console.log(`${"  ".repeat(depth)}- ${name} (${action.type})${runAfter ? "  after: " + runAfter : ""}`);
    checkAction(name, action);

    walk(action.actions, depth + 1);
    if (action.else) walk(action.else.actions, depth + 1);
    if (action.default) walk(action.default.actions, depth + 1);
    Object.keys(action.cases || {}).forEach((c) => walk(action.cases[c].actions, depth + 1));
  });
}

console.log(`\nFLOW: ${zipPath}\n`);
console.log("TRIGGER:");
Object.keys(flow.triggers || {}).forEach((name) => {
  const trigger = flow.triggers[name];
  console.log(`- ${name} (${trigger.type})`);
  scanValues(name, trigger.inputs);
  if (trigger.conditions) console.log("  conditions:", JSON.stringify(trigger.conditions));
});

console.log("\nACTIONS:");
walk(flow.actions, 0);

console.log("\nLIBRARIES named with getbytitle:", [...libraries].join(" | ") || "none");
console.log("SDG ACCOUNTS named in the flow:", [...accounts].join(" | ") || "none");
console.log(`\nWARNINGS (${warnings.length}):`);
[...new Set(warnings)].forEach((w) => console.log("- " + w));

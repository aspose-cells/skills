"use strict";

const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const requiredReferences = [
  "setup-and-licensing.md",
  "cross-platform-deployment.md",
  "cell-values-and-data.md",
  "styles-and-formatting.md",
  "formulas-and-calculation.md",
  "worksheets-rows-columns.md",
  "images-and-shapes.md",
  "charts.md",
  "pivot-tables.md",
  "conversion-and-rendering.md",
  "performance-and-large-files.md",
  "smart-markers-reporting.md",
  "vba-macros.md",
  "digital-signatures-and-document-properties.md",
  "slicers-and-sparklines.md",
  "ole-objects.md"
];

const errors = [];
let blockCount = 0;

function fail(file, message) {
  errors.push(`${path.relative(root, file) || path.basename(file)}: ${message}`);
}

function walk(directory) {
  const files = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...walk(fullPath));
    } else {
      files.push(fullPath);
    }
  }
  return files;
}

function validateFrontmatter(file, source) {
  const match = source.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
  if (!match) {
    fail(file, "missing YAML frontmatter");
    return;
  }

  const name = match[1].match(/^name:[ \t]*(.+?)[ \t]*$/m);
  const description = match[1].match(/^description:[ \t]*(.+?)[ \t]*$/m);
  if (!name || name[1] !== path.basename(root)) {
    fail(file, `name must equal ${path.basename(root)}`);
  }
  if (!description || !description[1].trim()) {
    fail(file, "description must be present and non-empty");
  }
}

function validateLinks(file, source) {
  const links = source.matchAll(/\[[^\]]*\]\(([^)]+)\)/g);
  for (const match of links) {
    let target = match[1].trim();
    if (/^(?:https?:|mailto:|#|file:|computer:)/i.test(target)) {
      continue;
    }
    target = target.split(/\s+["']/)[0];
    target = target.split("#")[0];
    if (!target) {
      continue;
    }
    try {
      target = decodeURIComponent(target);
    } catch {
      fail(file, `invalid encoded link ${match[1]}`);
      continue;
    }
    const resolved = path.resolve(path.dirname(file), target);
    if (!fs.existsSync(resolved)) {
      fail(file, `broken local link ${match[1]}`);
    }
  }
}

function validateJavaScript(file, source) {
  const blocks = source.matchAll(/```(?:javascript|js)[ \t]*\r?\n([\s\S]*?)```/gi);
  for (const match of blocks) {
    blockCount += 1;
    const line = source.slice(0, match.index).split(/\r?\n/).length + 1;
    try {
      new vm.Script(match[1], { filename: `${file}:${line}` });
    } catch (error) {
      fail(file, `JavaScript block at line ${line} does not parse: ${error.message}`);
    }
  }
}

const skillFile = path.join(root, "SKILL.md");
if (!fs.existsSync(skillFile)) {
  fail(skillFile, "required file is missing");
} else {
  const markdownFiles = walk(root).filter((file) => file.toLowerCase().endsWith(".md"));
  for (const file of markdownFiles) {
    const source = fs.readFileSync(file, "utf8");
    if (file === skillFile) {
      validateFrontmatter(file, source);
    }
    validateLinks(file, source);
    validateJavaScript(file, source);
  }

  for (const reference of requiredReferences) {
    const file = path.join(root, "references", reference);
    if (!fs.existsSync(file)) {
      fail(file, "required reference is missing");
    }
  }
}

if (errors.length > 0) {
  console.error(errors.join("\n"));
  process.exit(1);
}

console.log(`Static validation passed: ${requiredReferences.length} references, ${blockCount} JavaScript blocks.`);

if (process.argv.includes("--runtime")) {
  let temporaryDirectory;
  try {
    const cells = require("aspose.cells");
    temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "aspose-cells-nodejs-via-java-skill-"));

    if (process.env.ASPOSE_CELLS_LICENSE) {
      const license = new cells.License();
      license.setLicense(path.resolve(process.env.ASPOSE_CELLS_LICENSE));
    }

    const workbook = new cells.Workbook();
    const sheet = workbook.getWorksheets().get(0);
    sheet.getCells().get("A1").putValue("Price");
    sheet.getCells().get("B1").putValue(42.5);
    sheet.getCells().get("C1").setFormula("=B1*2");
    workbook.calculateFormula();

    const output = path.join(temporaryDirectory, "smoke.xlsx");
    workbook.save(output, cells.SaveFormat.XLSX);
    if (!fs.existsSync(output) || fs.statSync(output).size === 0) {
      throw new Error("The runtime smoke test did not create a non-empty workbook");
    }

    const reopened = new cells.Workbook(output);
    reopened.calculateFormula();
    const result = Number(reopened.getWorksheets().get(0).getCells().get("C1").getValue());
    if (result !== 85) {
      throw new Error(`Expected C1 to equal 85 after reopening, received ${result}`);
    }
    console.log("Runtime validation passed: workbook save/reopen and formula result are correct.");
  } catch (error) {
    console.error(`Runtime validation failed: ${error.stack || error.message}`);
    console.error("Install aspose.cells and its Java prerequisites, or run static validation without --runtime.");
    process.exitCode = 1;
  } finally {
    if (temporaryDirectory) {
      fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    }
  }
}

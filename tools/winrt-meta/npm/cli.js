#!/usr/bin/env node
const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const os = require("os");

const args = process.argv.slice(2);

// Parse --lang and --output from args
let lang = null;
let outputDir = null;
const exeArgs = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--lang") {
    lang = args[++i];
    exeArgs.push("--lang", "ts"); // exe always generates TS; shim handles js/cjs
  } else if (args[i] === "--output") {
    outputDir = args[++i];
    exeArgs.push("--output"); // placeholder, patched below
    exeArgs.push(null);
  } else {
    exeArgs.push(args[i]);
  }
}

const needsCompile = lang && lang !== "ts";

// If compiling, redirect exe output to a temp directory
let tsDir = outputDir;
if (needsCompile && outputDir) {
  tsDir = fs.mkdtempSync(path.join(os.tmpdir(), "winrt-meta-"));

  // Copy existing index.js as index.ts so exe can append to it (--class-name mode)
  const existingIndex = path.join(outputDir, "index.js");
  if (fs.existsSync(existingIndex)) {
    fs.copyFileSync(existingIndex, path.join(tsDir, "index.ts"));
  }
}

// Patch --output value in exeArgs
const outputIdx = exeArgs.indexOf(null);
if (outputIdx !== -1) {
  exeArgs[outputIdx] = tsDir || outputDir;
}

const arch = process.arch === "arm64" ? "arm64" : "x64";
const exe = path.join(__dirname, "bin", arch, "winrt-meta.exe");

// When compiling, suppress exe stdout (temp paths are noisy) but keep stderr
const stdio = needsCompile ? ["inherit", "pipe", "inherit"] : "inherit";

try {
  execFileSync(exe, exeArgs, { stdio });
} catch (e) {
  process.exit(e.status ?? 1);
}

if (needsCompile && outputDir) {
  const { compileDir } = require("./lib/compile");
  const moduleType = lang === "cjs" ? "commonjs" : "es6";

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  try {
    compileDir(tsDir, outputDir, { moduleType });
    console.log(`Done. Output in ${outputDir}`);
  } finally {
    fs.rmSync(tsDir, { recursive: true, force: true });
  }
}

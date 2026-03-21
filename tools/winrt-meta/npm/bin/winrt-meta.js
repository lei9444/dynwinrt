#!/usr/bin/env node
const { execFileSync } = require("child_process");
const path = require("path");

const arch = process.arch === "arm64" ? "arm64" : "x64";
const exe = path.join(__dirname, arch, "winrt-meta.exe");

try {
  execFileSync(exe, process.argv.slice(2), { stdio: "inherit" });
} catch (e) {
  process.exit(e.status ?? 1);
}

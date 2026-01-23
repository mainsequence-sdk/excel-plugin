#!/usr/bin/env node

/**
 * Wraps a command so NODE_OPTIONS always includes a persistent localStorage file.
 * Needed because Node 25 exposes browser storage APIs that require --localstorage-file.
 */

const { spawn } = require("child_process");

if (process.argv.length < 3) {
  console.error("Usage: node run-with-localstorage.js <command> [...args]");
  process.exit(1);
}

const FLAG = "--localstorage-file=.office-addin-localstorage";
const majorVersion = Number(process.versions.node?.split(".")[0] ?? 0);
const supportsLocalStorageFlag = Number.isFinite(majorVersion) && majorVersion >= 25;
const existingOptions = process.env.NODE_OPTIONS ? process.env.NODE_OPTIONS.split(" ") : [];

let nextOptions = existingOptions.filter(Boolean);

if (supportsLocalStorageFlag) {
  if (!nextOptions.includes(FLAG)) {
    nextOptions.push(FLAG);
  }
} else {
  // Older Node versions reject --localstorage-file, so strip it out entirely.
  nextOptions = nextOptions.filter((arg) => !arg.startsWith("--localstorage-file"));
}

process.env.NODE_OPTIONS = nextOptions.join(" ").trim();

const [command, ...args] = process.argv.slice(2);
const child = spawn(command, args, {
  stdio: "inherit",
  shell: process.platform === "win32",
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
  } else {
    process.exit(code ?? 0);
  }
});

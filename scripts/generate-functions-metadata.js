#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const { generateCustomFunctionsMetadata } = require("custom-functions-metadata");

const inputFile = path.resolve(__dirname, "../src/functions/functions.ts");
const outputFile = path.resolve(__dirname, "../functions.json");

async function main() {
  const result = await generateCustomFunctionsMetadata([inputFile], true);
  if (result.errors && result.errors.length > 0) {
    console.error("Failed to generate functions metadata:");
    result.errors.forEach((err) => console.error(`- ${err}`));
    process.exit(1);
  }

  fs.writeFileSync(outputFile, result.metadataJson);
  console.log(`Wrote ${path.relative(process.cwd(), outputFile)}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

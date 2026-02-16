#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (!arg.startsWith('--')) continue;
    const key = arg.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith('--')) {
      result[key] = true;
      continue;
    }
    result[key] = value;
    index += 1;
  }
  return result;
}

const args = parseArgs(process.argv.slice(2));

const sourcePath = path.resolve(args.source ?? 'manifest.xml');
const outputPath = path.resolve(args.output ?? 'manifests/manifest.staging.xml');
const baseUrlInput = args.baseUrl ?? 'https://sbroenne.github.io/office-coding-agent';
const addinId = args.id ?? '4f0b6a89-6cba-4930-b102-47c7dc4d9c2d';

const baseUrl = baseUrlInput.replace(/\/$/, '');

if (!fs.existsSync(sourcePath)) {
  throw new Error(`Source manifest not found: ${sourcePath}`);
}

let xml = fs.readFileSync(sourcePath, 'utf8');

xml = xml.replace(
  /<Id>[\s\S]*?<\/Id>/,
  `<Id>${addinId}</Id>`
);

xml = xml.replaceAll('https://localhost:3000', baseUrl);

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, xml, 'utf8');

console.log(`Generated staging manifest: ${outputPath}`);
console.log(`Base URL: ${baseUrl}`);
console.log(`Add-in ID: ${addinId}`);

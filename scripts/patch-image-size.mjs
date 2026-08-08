import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repositoryRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packageRoot = path.join(repositoryRoot, 'node_modules', 'image-size');

const patches = [
  {
    file: path.join(packageRoot, 'dist', 'types', 'icns.js'),
    vulnerable: `function readImageHeader(input, imageOffset) {
    const imageLengthOffset = imageOffset + ENTRY_LENGTH_OFFSET;
    return [
        (0, utils_1.toUTF8String)(input, imageOffset, imageLengthOffset),
        (0, utils_1.readUInt32BE)(input, imageLengthOffset),
    ];
}`,
    patched: `function readImageHeader(input, imageOffset) {
    const imageLengthOffset = imageOffset + ENTRY_LENGTH_OFFSET;
    const imageLength = (0, utils_1.readUInt32BE)(input, imageLengthOffset);
    if (!Number.isSafeInteger(imageLength) ||
        imageLength < SIZE_HEADER ||
        imageOffset + imageLength > input.length) {
        throw new TypeError('Invalid ICNS image entry length');
    }
    return [
        (0, utils_1.toUTF8String)(input, imageOffset, imageLengthOffset),
        imageLength,
    ];
}`,
  },
  {
    file: path.join(packageRoot, 'dist', 'types', 'utils.js'),
    vulnerable: `    const boxSize = (0, exports.readUInt32BE)(input, offset);
    if (input.length - offset < boxSize)
        return;
    return {
        name: (0, exports.toUTF8String)(input, 4 + offset, 8 + offset),
        offset,
        size: boxSize,
    };`,
    patched: `    const declaredBoxSize = (0, exports.readUInt32BE)(input, offset);
    const boxSize = declaredBoxSize === 0 ? input.length - offset : declaredBoxSize;
    if (boxSize < 8 || input.length - offset < boxSize)
        return;
    return {
        name: (0, exports.toUTF8String)(input, 4 + offset, 8 + offset),
        offset,
        size: boxSize,
    };`,
  },
];

function readPackageVersion() {
  const manifestPath = path.join(packageRoot, 'package.json');
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
  if (manifest.name !== 'image-size' || manifest.version !== '1.2.1') {
    throw new Error(
      `Unsupported image-size package ${manifest.name}@${manifest.version}; review the security patch before updating.`,
    );
  }
  return manifest.version;
}

export function verifyImageSizeSecurityPatch() {
  readPackageVersion();
  const failures = [];
  for (const patch of patches) {
    const source = readFileSync(patch.file, 'utf8');
    if (!source.includes(patch.patched)) {
      failures.push(path.relative(repositoryRoot, patch.file));
    }
  }
  return failures;
}

export function applyImageSizeSecurityPatch() {
  const version = readPackageVersion();

  for (const patch of patches) {
    const source = readFileSync(patch.file, 'utf8');
    if (source.includes(patch.patched)) {
      continue;
    }
    if (!source.includes(patch.vulnerable)) {
      throw new Error(
        `Cannot safely patch ${path.relative(repositoryRoot, patch.file)} because its source is unexpected.`,
      );
    }
    writeFileSync(patch.file, source.replace(patch.vulnerable, patch.patched), 'utf8');
  }

  const failures = verifyImageSizeSecurityPatch();
  if (failures.length > 0) {
    throw new Error(`image-size security patch verification failed for: ${failures.join(', ')}`);
  }

  console.log(`Applied verified image-size@${version} parser safeguards.`);
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : '';
if (invokedPath === fileURLToPath(import.meta.url)) {
  applyImageSizeSecurityPatch();
}

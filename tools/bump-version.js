const fs = require('fs');
const path = require('path');

const validActions = ['major', 'minor', 'patch', 'sync'];
const action = process.argv[2];

function logUsage() {
  console.error('Usage: node tools/bump-version.js <major|minor|patch|sync>');
  process.exit(1);
}

if (!action || !validActions.includes(action)) {
  logUsage();
}

const root = path.resolve(__dirname, '..');
const packageJsonPath = path.join(root, 'package.json');
const solutionJsonPath = path.join(root, 'config', 'package-solution.json');

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function writeJson(filePath, data) {
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2) + '\n', 'utf8');
}

function parseSemver(version) {
  const parts = version.split('.');
  if (parts.length !== 3 || parts.some((part) => !/^[0-9]+$/.test(part))) {
    throw new Error(`Invalid semver version: ${version}`);
  }
  return parts.map(Number);
}

function parseSolutionVersion(version) {
  const parts = version.split('.');
  if (parts.length !== 4 || parts.some((part) => !/^[0-9]+$/.test(part))) {
    throw new Error(`Invalid solution version: ${version}`);
  }
  return parts.map(Number);
}

function bumpVersion(major, minor, patch) {
  if (action === 'major') {
    major += 1;
    minor = 0;
    patch = 0;
  } else if (action === 'minor') {
    minor += 1;
    patch = 0;
  } else if (action === 'patch') {
    patch += 1;
  }
  return [major, minor, patch];
}

function formatSemver(parts) {
  return parts.join('.');
}

function formatSolutionVersion(parts) {
  return `${parts[0]}.${parts[1]}.${parts[2]}.${parts[3]}`;
}

function updateFeatureVersions(features, major, minor, patch) {
  if (!Array.isArray(features)) {
    return features;
  }

  return features.map((feature) => {
    if (feature && typeof feature.version === 'string') {
      const parts = feature.version.split('.');
      if (parts.length === 4 && parts.slice(0, 3).every((part) => /^[0-9]+$/.test(part))) {
        const last = /^[0-9]+$/.test(parts[3]) ? Number(parts[3]) : 0;
        feature.version = formatSolutionVersion([major, minor, patch, last]);
      }
    }
    return feature;
  });
}

const packageJson = readJson(packageJsonPath);
const oldPackageVersion = packageJson.version;
let [major, minor, patch] = parseSemver(oldPackageVersion);

if (action !== 'sync') {
  [major, minor, patch] = bumpVersion(major, minor, patch);
  const newPackageVersion = formatSemver([major, minor, patch]);
  packageJson.version = newPackageVersion;
  writeJson(packageJsonPath, packageJson);
  console.log(`Updated package.json version: ${oldPackageVersion} -> ${newPackageVersion}`);
} else {
  console.log(`Sync mode: using existing package.json version ${oldPackageVersion}`);
}

const newSolutionVersion = `${major}.${minor}.${patch}.0`;
const solutionJson = readJson(solutionJsonPath);
const oldSolutionVersion = solutionJson.solution?.version;

if (!oldSolutionVersion) {
  throw new Error('package-solution.json does not contain solution.version');
}

solutionJson.solution.version = newSolutionVersion;
solutionJson.solution.features = updateFeatureVersions(solutionJson.solution.features, major, minor, patch);
writeJson(solutionJsonPath, solutionJson);
console.log(`Updated config/package-solution.json solution.version: ${oldSolutionVersion} -> ${newSolutionVersion}`);

if (Array.isArray(solutionJson.solution.features)) {
  solutionJson.solution.features.forEach((feature) => {
    if (feature && typeof feature.version === 'string') {
      console.log(`Updated feature version: ${feature.title || 'feature'} -> ${feature.version}`);
    }
  });
}

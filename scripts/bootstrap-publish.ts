import {spawnSync} from 'node:child_process';
import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const registry = 'https://registry.npmjs.org/';
const manifest: unknown = JSON.parse(
  readFileSync(new URL('../package.json', import.meta.url), 'utf8'),
);
if (
  typeof manifest !== 'object' ||
  manifest === null ||
  !('name' in manifest) ||
  !('version' in manifest) ||
  typeof manifest.name !== 'string' ||
  typeof manifest.version !== 'string'
) {
  throw new Error('package.json must contain a package name and version');
}
const {name, version} = manifest;

function run(command: string, args: string[], capture = false): string {
  const result = spawnSync(command, args, {
    cwd: root,
    stdio: capture ? ['inherit', 'pipe', 'inherit'] : 'inherit',
    encoding: 'utf8',
    shell: process.platform === 'win32',
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} ${args[0]} failed (${result.status})`);
  return result.stdout?.trim() ?? '';
}

if (!process.stdin.isTTY || !process.stdout.isTTY) {
  throw new Error('Run this script in an interactive terminal for npm login and 2FA.');
}
if (name !== '@lofcz/ts-xlsx') throw new Error('Unexpected package: refusing to publish.');
const npmVersion = run('npm', ['--version'], true);
const [major, minor] = npmVersion.split('.').map(Number);
if (
  major === undefined ||
  minor === undefined ||
  !Number.isFinite(major) ||
  major < 11 ||
  (major === 11 && minor < 15)
) {
  throw new Error('npm trust requires npm >= 11.15.0. Update npm, then rerun.');
}

process.stdout.write(
  `First publish: ${name}@${version}; trusted publisher: lofcz/ts-xlsx, release.yml\n`,
);
run('pnpm', ['install', '--frozen-lockfile']);
// prepublishOnly runs verification, including the build/size gate, and the distribution smoke.
run('npm', ['login', '--auth-type=web', `--registry=${registry}`]);
const response = await fetch(`${registry}${encodeURIComponent(name)}/${version}`);
if (response.status === 404) {
  run('npm', ['publish', '--access=public', `--registry=${registry}`]);
} else if (response.ok) {
  process.stdout.write(`${name}@${version} already exists; continuing with trust configuration.\n`);
} else {
  throw new Error(`Cannot check published version: HTTP ${response.status}`);
}

// Trust configuration is an account-authenticated, interactive 2FA operation.
// Do not accept a token/OTP argument or write credentials to the repository.
run('npm', ['trust', 'list', name, `--registry=${registry}`]);
run('npm', [
  'trust',
  'github',
  name,
  '--repo=lofcz/ts-xlsx',
  '--file=release.yml',
  '--allow-publish',
  `--registry=${registry}`,
]);
run('npm', ['trust', 'list', name, `--registry=${registry}`]);
process.stdout.write(
  'Bootstrap complete. Future releases: gh workflow run release.yml --repo lofcz/ts-xlsx -f release_type=patch\n',
);

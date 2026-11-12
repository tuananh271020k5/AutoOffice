import { spawn } from 'node:child_process';

// Preserve the project's normal sideload flow, including automatic Vite startup.
const child = spawn(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['run', 'sideload'], {
  stdio: 'inherit',
  shell: process.platform === 'win32',
  env: { ...process.env, AUTOOFFICE_TEST_BRIDGE: '1' },
});
child.on('error', error => { console.error(error.message); process.exitCode = 1; });
child.on('exit', code => { process.exitCode = code ?? 1; });

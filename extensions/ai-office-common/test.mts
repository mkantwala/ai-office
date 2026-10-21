/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
const { build } = createRequire(new URL('../../build/package.json', import.meta.url))('esbuild');
await build({ entryPoints: ['extensions/ai-office-common/engine.test.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/office-engine.test.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/provider.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode', './client.js', './agentRequest.js'], outfile: '.build/office-provider.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/cloudflareClient.test.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/office-cloudflareClient.test.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/wordFormatting.test.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/office-wordFormatting.test.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/documentCapabilities.test.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/office-capabilities.test.cjs' });
const filter = process.argv.includes('--word') ? ['--test-name-pattern=^(Word|Cloudflare|custom editor)'] : [];
const result = spawnSync(process.execPath, ['--test', ...filter, '.build/office-engine.test.cjs', '.build/office-cloudflareClient.test.cjs', '.build/office-wordFormatting.test.cjs', '.build/office-capabilities.test.cjs', 'extensions/ai-office-common/provider.test.cts'], { stdio: 'inherit' });
process.exitCode = result.status ?? 1;

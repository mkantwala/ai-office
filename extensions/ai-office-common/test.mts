/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { build } from 'esbuild';
import { spawnSync } from 'node:child_process';
await build({ entryPoints: ['extensions/ai-office-common/engine.test.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/office-engine.test.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/provider.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode', './client.js'], outfile: '.build/office-provider.cjs' });
await build({ entryPoints: ['extensions/ai-office-common/wordProvider.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode', './provider.js'], outfile: '.build/office-word-provider.cjs' });
const result = spawnSync(process.execPath, ['--test', '.build/office-engine.test.cjs', 'extensions/ai-office-common/provider.test.cts', 'extensions/ai-office-common/wordProvider.test.cts'], { stdio: 'inherit' });
process.exitCode = result.status ?? 1;

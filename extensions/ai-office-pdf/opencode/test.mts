/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { build } from 'esbuild';
import { spawnSync } from 'node:child_process';
await build({ entryPoints: ['extensions/ai-office-pdf/opencode/test/transport.test.mts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/opencode-transport.test.cjs' });
await build({ entryPoints: ['extensions/ai-office-pdf/opencode/provider.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode'], outfile: '.build/opencode-provider-test.cjs' });
const result = spawnSync(process.execPath, ['--test', '.build/opencode-transport.test.cjs', 'extensions/ai-office-pdf/opencode/test/provider.test.cts'], { stdio: 'inherit' });
process.exitCode = result.status ?? 1;

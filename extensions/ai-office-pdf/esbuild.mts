/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const result = spawnSync(process.execPath, [fileURLToPath(new URL('./build.mts', import.meta.url)), '--ui-only'], { stdio: 'inherit' });
if (result.error) { throw result.error; }
process.exitCode = result.status ?? 1;

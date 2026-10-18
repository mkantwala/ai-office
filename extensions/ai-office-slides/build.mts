/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { build } from 'esbuild';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
const root = dirname(fileURLToPath(import.meta.url));
await Promise.all([
	build({ entryPoints: [join(root, '../ai-office-common/engineWorker.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(root, 'out/engineWorker.cjs') }),
	build({ entryPoints: [join(root, 'extension.ts')], outfile: join(root, 'out/extension.js'), bundle: true, platform: 'node', format: 'cjs', external: ['vscode'], sourcemap: true }),
	build({ entryPoints: [join(root, '../ai-office-common/editor.ts')], outfile: join(root, 'out/editor.js'), bundle: true, platform: 'browser', sourcemap: true })
]);
await writeFile(join(root, 'out/editor.css'), (await readFile(join(root, '../../src/vs/base/browser/ui/documentToolbar/documentToolbar.css'), 'utf8')) + '\n' + (await readFile(join(root, '../ai-office-common/editor.css'), 'utf8')));
await mkdir(join(root, 'out/licenses'), { recursive: true });
for (const [name, file] of [['pptx', 'node_modules/@office-kit/pptx/LICENSE'], ['pptx-preview', 'node_modules/@office-kit/pptx-preview/LICENSE'], ['dompurify', 'node_modules/dompurify/LICENSE'], ['fflate', 'node_modules/fflate/LICENSE']]) {
	await copyFile(join(root, file), join(root, `out/licenses/${name}.txt`));
}

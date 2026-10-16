/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import fs from 'fs';
import path from 'path';

const root = path.resolve(import.meta.dirname, '../..');
const product: { disabledBuiltinExtensions?: string[] } = JSON.parse(fs.readFileSync(path.join(root, 'product.json'), 'utf8'));
const disabledExtensions = new Set(product.disabledBuiltinExtensions?.map(id => id.toLowerCase()));

export function isBuiltinExtensionEnabled(manifest: { publisher?: string; name: string }): boolean {
	return !disabledExtensions.has(`${manifest.publisher}.${manifest.name}`.toLowerCase());
}

/** Apply the product policy to nested TypeScript, browser and media build inputs. */
export function isBuiltinExtensionBuildEnabled(buildPath: string): boolean {
	const [directory, extension] = path.relative(root, path.resolve(root, buildPath)).split(path.sep);
	if (directory !== 'extensions') {
		return true; // Repository development tools are not shipped built-in extensions.
	}
	const manifest = JSON.parse(fs.readFileSync(path.join(root, 'extensions', extension, 'package.json'), 'utf8'));
	return isBuiltinExtensionEnabled(manifest);
}

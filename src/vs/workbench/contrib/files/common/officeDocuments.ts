/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** File types shown in the AI Office Explorer. Folder visibility is handled separately. */
export function isOfficeDocument(name: string): boolean {
	return /\.(docx|xlsx|pptx|pdf|md|markdown)$/i.test(name);
}

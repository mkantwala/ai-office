/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export interface IOfficeDocumentReview {
	uri: string;
	version: string;
	reviewVersion: number;
	changes: { id: string; label: string; before: string; after: string }[];
}

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export interface PageMargins { top: number; bottom: number; left: number; right: number; gutter: number }
export interface PageBorder { style: 'single' | 'double' | 'dotted' | 'dashed' | 'thick' | 'triple' | 'dotDash' | 'wave'; width: number; color: string | null; space: number }
export type PageBorderSides = Partial<Record<'top' | 'bottom' | 'left' | 'right', PageBorder | null>>;
export interface PageSetupState { page: number; section: number; width: number; height: number; margins: PageMargins; borders: PageBorderSides }
export interface PagePropertyPatch { page: number; applyTo: 'section' | 'document'; margins?: Partial<PageMargins>; borders?: PageBorderSides | null }

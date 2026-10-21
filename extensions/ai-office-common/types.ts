/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { PageSetupState } from './pageProperties.js';
import type { ParagraphProperties } from './paragraphProperties.js';

export type Kind = 'word' | 'pdf' | 'slides';
export interface Block { id: number; text: string; rect: number[]; font?: string; size?: number; row?: number; col?: number }
export interface TableGeometry { cols: number[]; rows: { row: number; top: number; bottom: number; edges?: number[] }[]; endsHere: boolean; movable?: boolean }
export interface TablePreview { kind: 'column' | 'row' | 'corner' | 'move'; rects: { x: number; y: number; width: number; height: number }[] }
export interface ImageGeometry {
	parentId: string; offset: number; page: number; x: number; y: number; width: number; height: number;
	alt: string; title: string; decorative: boolean; crop: number[]; wrap: string; rotation: number; flipH: boolean; flipV: boolean;
}
export interface State {
	image?: ImageGeometry | null;
	pageSetup?: PageSetupState;
	editingRegion?: 'body' | 'header' | 'footer';
	pageRegions?: { top: number; bottom: number; width: number; height: number }[];
	formatting?: { font: string | null; size: number | null; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean };
	paragraph?: { id: string; properties: ParagraphProperties; mixed: string[]; ruler?: { left: number; width: number } | null };
	pageRuler?: { left: number; width: number } | null;
	undoDepth?: number;
	agentAppliedEdits?: { id: string; operation: string; command?: string; params?: unknown }[];
	agentCreatedBlocks?: { id: string; kind: string; parentId: string }[];
	agentCommandResult?: unknown;
	revision: number; undo?: boolean; redo?: boolean; page?: number; blocks?: Block[]; matches?: Block[];
	caret?: { page: number; x: number; y: number; height: number; width?: number };
	rects?: { page: number; x: number; y: number; width: number; height: number }[];
	selection?: unknown; selectedText?: string;
	paragraphStyle?: string;
	comments?: { commentId?: number; parentId?: number | null; resolved?: boolean; author?: string; text: string }[];
}
export interface Command { id: string; label: string; params?: unknown }
export interface DocumentAgent {
	read(args: Record<string, unknown>): unknown;
	edit(args: Record<string, unknown>): Promise<State>;
	rects(id: string): State['rects'];
	reviewStart?(): void;
	reviewFinish?(): void;
	reviewRead?(): DocumentReviewChanges;
	reviewDecide?(revision: number, action: 'keep' | 'undo', ids: string[]): State;
	reviewExport?(): string;
	reviewImport?(value: string): void;
	reviewProjection?(): string;
}
export interface DraftViewSection { id: string; title: string; intent: string; anchor: { id: string; position: 'before' | 'after' }; anchorRects?: State['rects'] }
export interface DocumentReviewChange {
	id: string; label: string; before: string; after: string;
	groupId?: string;
	beforeText?: string; afterText?: string;
	anchor?: { id: string; after: boolean };
	rects?: NonNullable<State['rects']>;
	anchorRects?: NonNullable<State['rects']>;
}
export interface DocumentReviewChanges { changes: DocumentReviewChange[] }
export interface DocumentReview extends DocumentReviewChanges { uri: string; version: string; reviewVersion: number }
export interface Engine {
	tablePreview?(args: Record<string, unknown>): TablePreview;
	agent?: DocumentAgent;
	pages(): number[][];
	state(): State;
	commands(): Command[];
	execute(id: string, args: Record<string, unknown>): Promise<State>;
	export(): Promise<Uint8Array>;
	render(page: number, scale: number): { pixels?: Uint8Array; width?: number; height?: number; svg?: string; tables?: TableGeometry[]; images?: ImageGeometry[] };
	dispose(): void;
}

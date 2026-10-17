/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export type Kind = 'word' | 'pdf' | 'slides';
export interface Block { id: number; text: string; rect: number[]; font?: string; size?: number; row?: number; col?: number }
export interface State {
	agentCreatedBlocks?: { id: string; kind: string; parentId: string }[];
	agentCommandResult?: unknown;
	revision: number; undo?: boolean; redo?: boolean; page?: number; blocks?: Block[];
	caret?: { page: number; x: number; y: number; height: number };
	rects?: { page: number; x: number; y: number; width: number; height: number }[];
	selection?: unknown; selectedText?: string;
	paragraphStyle?: string;
	comments?: { author?: string; text: string }[];
}
export interface Command { id: string; label: string; params?: unknown }
export interface DocumentAgent {
	read(args: Record<string, unknown>): unknown;
	edit(args: Record<string, unknown>): Promise<State>;
	rects(id: string): State['rects'];
}
export interface Engine {
	agent?: DocumentAgent;
	pages(): number[][];
	state(): State;
	commands(): Command[];
	execute(id: string, args: Record<string, unknown>): Promise<State>;
	export(): Promise<Uint8Array>;
	render(page: number, scale: number): { pixels?: Uint8Array; width?: number; height?: number; svg?: string };
	dispose(): void;
}

/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import '../../../../base/browser/ui/documentToolbar/documentToolbar.css';
import { createDocumentToolbar, isDocumentToolbarSnapshot, type DocumentToolbarSnapshot, type DocumentToolbarAction } from '../../../../base/browser/ui/documentToolbar/documentToolbar.js';
import { URI } from '../../../../base/common/uri.js';
import { Disposable, MutableDisposable } from '../../../../base/common/lifecycle.js';
import { ICommandService, CommandsRegistry } from '../../../../platform/commands/common/commands.js';
import { INotificationService } from '../../../../platform/notification/common/notification.js';
import { localize } from '../../../../nls.js';
import { IWorkbenchContribution, WorkbenchPhase, registerWorkbenchContribution2 } from '../../../common/contributions.js';
import { IEditorService } from '../../../services/editor/common/editorService.js';
import { IEditorGroupsService } from '../../../services/editor/common/editorGroupsService.js';
import { CustomEditorInput } from '../../customEditor/browser/customEditorInput.js';

/** One window-level control surface; document selection remains owned by its editor. */
class DocumentToolbarContribution extends Disposable implements IWorkbenchContribution {
	static readonly ID = 'workbench.contrib.documentToolbar';
	private readonly registration = this._register(new MutableDisposable());
	private readonly host = document.createElement('div');
	private readonly toolbar = document.createElement('header');
	private controls?: ReturnType<typeof createDocumentToolbar>;
	private target?: DocumentToolbarSnapshot;
	private interaction?: DocumentToolbarSnapshot;
	private shape = '';
	private input?: CustomEditorInput;
	private ready = false;

	constructor(
		@IEditorService private readonly editors: IEditorService,
		@IEditorGroupsService private readonly groups: IEditorGroupsService,
		@ICommandService private readonly commands: ICommandService,
		@INotificationService private readonly notifications: INotificationService,
	) {
		super();
		this.host.className = 'sangam-document-toolbar'; this.host.append(this.toolbar);
		this._register(CommandsRegistry.registerCommand('_sangamOne.documentToolbar.supported', () => this.groups.getPart(this.groups.activeGroup) === this.groups.mainPart));
		this._register(CommandsRegistry.registerCommand('_sangamOne.documentToolbar.update', (_accessor, resource: unknown, payload: unknown) => {
			if (!URI.isUri(resource) || !payload || typeof payload !== 'object') { return false; }
			const value = { ...payload, uri: resource.toString() };
			if (!isDocumentToolbarSnapshot(value)) { return false; }
			const input = this.activeInput();
			if (!input || input.resource.toString() !== value.uri || input.viewType !== `aiOffice.${value.kind}`) { return false; }
			this.target = value;
			const shape = JSON.stringify([value.kind, value.controls.map(({ key, type, label, text, options }) => ({ key, type, label, text, options }))]);
			if (shape !== this.shape) {
				this.controls?.dispose(); this.toolbar.replaceChildren(); this.shape = shape;
				this.controls = createDocumentToolbar(this.toolbar, value.controls, action => this.dispatch(action), () => this.editors.activeEditorPane?.focus());
			}
			this.controls?.update(value.controls, value.paragraphStyle);
			this.host.setAttribute('aria-busy', 'false');
			return true;
		}));
		const capture = () => { this.interaction = this.target; };
		this.host.addEventListener('pointerdown', capture); this.host.addEventListener('focusin', capture);
		this._register({ dispose: () => { this.host.removeEventListener('pointerdown', capture); this.host.removeEventListener('focusin', capture); this.controls?.dispose(); this.host.remove(); } });
		this._register(this.editors.onDidActiveEditorChange(() => this.activeChanged()));
		void this.groups.whenReady.then(() => { this.ready = true; this.activeChanged(); });
	}

	private activeInput(): CustomEditorInput | undefined {
		const input = this.editors.activeEditor;
		return this.groups.getPart(this.groups.activeGroup) === this.groups.mainPart && input instanceof CustomEditorInput && ['aiOffice.word', 'aiOffice.pdf', 'aiOffice.slides'].includes(input.viewType) ? input : undefined;
	}

	private activeChanged() {
		if (!this.ready) { return; }
		const input = this.activeInput(); if (input === this.input) { return; } this.input = input;
		this.target = undefined; this.interaction = undefined; this.shape = '';
		this.controls?.dispose(); this.controls = undefined; this.toolbar.replaceChildren();
		if (!input) { this.registration.clear(); return; }
		if (!this.registration.value) { this.registration.value = this.groups.mainPart.registerToolbar(this.host, 48); }
		this.host.setAttribute('aria-busy', 'true');
		this.toolbar.setAttribute('aria-label', localize('documentToolbar.loading', "Loading Document Formatting"));
		void this.commands.executeCommand(`_aiOffice.${input.viewType.slice('aiOffice.'.length)}.requestToolbar`, input.resource).catch(() => { /* The document extension publishes its controls when ready. */ });
	}

	private dispatch(action: DocumentToolbarAction) {
		const target = this.interaction || this.target; this.interaction = undefined;
		if (!target || target.context !== this.target?.context || this.activeInput()?.resource.toString() !== target.uri || this.activeInput()?.viewType !== `aiOffice.${target.kind}`) { return; }
		void this.commands.executeCommand<boolean>(`_aiOffice.${target.kind}.toolbarAction`, URI.parse(target.uri), target.panel, target.instance, target.context, action).then(applied => {
			if (applied && this.activeInput()?.resource.toString() === target.uri) { this.editors.activeEditorPane?.focus(); }
		}, error => this.notifications.error(error));
	}
}

registerWorkbenchContribution2(DocumentToolbarContribution.ID, DocumentToolbarContribution, WorkbenchPhase.BlockRestore);

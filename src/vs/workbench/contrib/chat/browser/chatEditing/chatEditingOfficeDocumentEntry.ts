/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { constObservable, IObservable } from '../../../../../base/common/observable.js';
import { assertType } from '../../../../../base/common/types.js';
import { URI } from '../../../../../base/common/uri.js';
import { Schemas } from '../../../../../base/common/network.js';
import { toLocalResource } from '../../../../../base/common/resources.js';
import { getCodeEditor } from '../../../../../editor/browser/editorBrowser.js';
import { IDocumentDiff } from '../../../../../editor/common/diff/documentDiffProvider.js';
import { DetailedLineRangeMapping } from '../../../../../editor/common/diff/rangeMapping.js';
import { LineRange } from '../../../../../editor/common/core/ranges/lineRange.js';
import { ILanguageService } from '../../../../../editor/common/languages/language.js';
import { ITextModel } from '../../../../../editor/common/model.js';
import { IModelService } from '../../../../../editor/common/services/model.js';
import { ITextModelService } from '../../../../../editor/common/services/resolverService.js';
import { ICommandService } from '../../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { IFileService } from '../../../../../platform/files/common/files.js';
import { IInstantiationService } from '../../../../../platform/instantiation/common/instantiation.js';
import { IUndoRedoService } from '../../../../../platform/undoRedo/common/undoRedo.js';
import { IAgentHostConnectionsService } from '../../../../../platform/agentHost/common/agentHostConnectionsService.js';
import { IEditorPane } from '../../../../common/editor.js';
import { IFilesConfigurationService } from '../../../../services/filesConfiguration/common/filesConfigurationService.js';
import { IWorkbenchEnvironmentService } from '../../../../services/environment/common/environmentService.js';
import { ICustomEditorService } from '../../../customEditor/common/customEditor.js';
import { IAiEditTelemetryService } from '../../../editTelemetry/browser/telemetry/aiEditTelemetry/aiEditTelemetryService.js';
import { IChatService } from '../../common/chatService/chatService.js';
import { ChatEditKind, IModifiedEntryTelemetryInfo, IModifiedFileEntryEditorIntegration, ISnapshotEntry, ModifiedFileEntryState } from '../../common/editing/chatEditingService.js';
import { IOfficeDocumentReview } from '../../common/editing/officeDocumentReview.js';
import { IChatResponseModel } from '../../common/model/chatModel.js';
import { AbstractChatEditingModifiedFileEntry } from './chatEditingModifiedFileEntry.js';
import { ChatEditingCodeEditorIntegration, IDocumentDiff2 } from './chatEditingCodeEditorIntegration.js';
import { ChatEditingTextModelChangeService } from './chatEditingTextModelChangeService.js';
import { ChatEditingTextModelContentProvider } from './chatEditingTextModelContentProviders.js';

/** Native chat review over a readable document projection; DOCX bytes remain in the document engine. */
export class ChatEditingOfficeDocumentEntry extends AbstractChatEditingModifiedFileEntry {
	static readonly reviewScheme = 'ai-office-review';
	static resource(uri: string) { return URI.from({ scheme: this.reviewScheme, path: URI.parse(uri).path + '.review.md', query: uri }); }
	readonly originalURI: URI;
	readonly originalModel: ITextModel;
	readonly modifiedModel: ITextModel;
	initialContent = '';
	private readonly changeService: ChatEditingTextModelChangeService;
	readonly diffInfo: IObservable<IDocumentDiff2>;
	private spans: { id: string; originalStart: number; originalEnd: number; modifiedStart: number; modifiedEnd: number }[] = [];
	private refresh: Promise<void> = Promise.resolve();
	private deciding = false;
	private readonly documentReference: Promise<void>;

	constructor(
		private data: IOfficeDocumentReview,
		telemetry: IModifiedEntryTelemetryInfo,
		@IModelService models: IModelService,
		@ILanguageService languages: ILanguageService,
		@ITextModelService resolver: ITextModelService,
		@ICommandService private readonly commands: ICommandService,
		@IConfigurationService config: IConfigurationService,
		@IFilesConfigurationService filesConfig: IFilesConfigurationService,
		@IChatService chat: IChatService,
		@IFileService files: IFileService,
		@IUndoRedoService undoRedo: IUndoRedoService,
		@IInstantiationService instantiation: IInstantiationService,
		@IAiEditTelemetryService telemetryService: IAiEditTelemetryService,
		@IAgentHostConnectionsService connections: IAgentHostConnectionsService,
		@ICustomEditorService customEditors: ICustomEditorService,
		@IWorkbenchEnvironmentService environment: IWorkbenchEnvironmentService,
	) {
		super(ChatEditingOfficeDocumentEntry.resource(data.uri), telemetry, ChatEditKind.Modified, config, filesConfig, chat, files, undoRedo, instantiation, telemetryService, connections);
		// Closing the Word tab must not dispose the engine while its changes are
		// still being reviewed in the diff editor.
		const documentURI = URI.parse(data.uri);
		const resource = documentURI.scheme === Schemas.file && environment.remoteAuthority ? toLocalResource(documentURI, environment.remoteAuthority, Schemas.vscodeRemote) : documentURI;
		const reference = customEditors.models.tryRetain(resource, 'aiOffice.word');
		this.documentReference = reference ? reference.then(value => {
			if (this._store.isDisposed) { value.dispose(); }
			else { this._register(value); }
		}) : Promise.resolve();
		this.originalURI = ChatEditingTextModelContentProvider.getFileURI(telemetry.sessionResource, this.entryId, this.modifiedURI.path);
		this.originalModel = this._register(models.createModel('', languages.createById('markdown'), this.originalURI));
		this.modifiedModel = this._register(models.createModel('', languages.createById('markdown'), this.modifiedURI));
		this._register(resolver.registerTextModelContentProvider(ChatEditingOfficeDocumentEntry.reviewScheme, { provideTextContent: async uri => models.getModel(uri) }));
		this.changeService = this._register(instantiation.createInstance(ChatEditingTextModelChangeService, this.originalModel, this.modifiedModel, this.state, undefined));
		this.diffInfo = this.changeService.diffInfo.map(diff => {
			// One decision covers the entire paragraph/table. Keep text and formatting
			// visible together instead of silently accepting a second, separate hunk.
			const groups: { ids: Set<string>; changes: DetailedLineRangeMapping[] }[] = [];
			for (const change of diff.changes) {
				const ids = new Set(this.idsForChange(change));
				const related = groups.filter(group => [...ids].some(id => group.ids.has(id)));
				const changes = [change];
				for (const group of related) {
					for (const id of group.ids) { ids.add(id); }
					changes.push(...group.changes); groups.splice(groups.indexOf(group), 1);
				}
				groups.push({ ids, changes });
			}
			for (const group of groups) { group.changes.sort((a, b) => a.original.startLineNumber - b.original.startLineNumber); }
			const changes = groups.map(group => new DetailedLineRangeMapping(
				new LineRange(Math.min(...group.changes.map(change => change.original.startLineNumber)), Math.max(...group.changes.map(change => change.original.endLineNumberExclusive))),
				new LineRange(Math.min(...group.changes.map(change => change.modified.startLineNumber)), Math.max(...group.changes.map(change => change.modified.endLineNumberExclusive))),
				group.changes.flatMap(change => change.innerChanges || [])
			)).sort((a, b) => a.modified.startLineNumber - b.modified.startLineNumber);
			return { ...diff, changes, keep: (change: DetailedLineRangeMapping) => this.decide('keep', change), undo: (change: DetailedLineRangeMapping) => this.decide('undo', change) };
		});
		this.update(data);
	}

	get changesCount() { return this.diffInfo.map(diff => diff.changes.length); }
	get linesAdded() { return this.changeService.diffInfo.map(diff => diff.changes.reduce((sum, change) => sum + change.modified.length, 0)); }
	get linesRemoved() { return this.changeService.diffInfo.map(diff => diff.changes.reduce((sum, change) => sum + change.original.length, 0)); }

	update(data: IOfficeDocumentReview) {
		this.data = data;
		this.refresh = this.refresh.then(async () => {
			let original = '', modified = '';
			const spans: typeof this.spans = [];
			for (const [index, change] of data.changes.entries()) {
				const header = `### ${change.label} ${index + 1}\n\n`;
				const originalStart = original.split('\n').length, modifiedStart = modified.split('\n').length;
				original += header + change.before + '\n\n---\n\n';
				modified += header + change.after + '\n\n---\n\n';
				spans.push({ id: change.id, originalStart, originalEnd: original.split('\n').length, modifiedStart, modifiedEnd: modified.split('\n').length });
			}
			this.spans = spans;
			this._stateObs.set(data.changes.length ? ModifiedFileEntryState.Modified : ModifiedFileEntryState.Accepted, undefined);
			await this.changeService.resetDocumentValues(original, modified);
			this.initialContent = original;
		});
		return this.refresh;
	}
	private idsForChange(change: DetailedLineRangeMapping) {
		return this.spans.filter(span =>
			(change.original.startLineNumber < span.originalEnd && change.original.endLineNumberExclusive > span.originalStart) ||
			(change.modified.startLineNumber < span.modifiedEnd && change.modified.endLineNumberExclusive > span.modifiedStart)
		).map(span => span.id);
	}

	private async decide(action: 'keep' | 'undo', change?: DetailedLineRangeMapping): Promise<boolean> {
		await this.documentReference;
		await this.refresh;
		if (this.deciding) { return false; }
		await this.changeService.getDiffInfo();
		const diff = this.diffInfo.get();
		if (change && !diff.changes.includes(change)) { return false; }
		let ids = change ? this.idsForChange(change) : this.data.changes.map(item => item.id);
		this.deciding = true;
		try {
			// Typing in Word must stay fast. Refresh at the decision boundary rather
			// than recomputing review projections after every user keystroke.
			const current = await this.commands.executeCommand<IOfficeDocumentReview>('aiOffice.word.reviewStatus', this.data.uri);
			assertType(current);
			if (JSON.stringify(current.changes) !== JSON.stringify(this.data.changes)) { await this.update(current); }
			else { this.data = current; }
			ids = change ? ids.filter(id => current.changes.some(item => item.id === id)) : current.changes.map(item => item.id);
			if (!ids.length) { return false; }
			const data = await this.commands.executeCommand<IOfficeDocumentReview>('aiOffice.word.reviewDecide', this.data.uri, this.data.version, this.data.reviewVersion, action, ids);
			assertType(data);
			await this.update(data);
			if (!data.changes.length) { this._stateObs.set(action === 'keep' ? ModifiedFileEntryState.Accepted : ModifiedFileEntryState.Rejected, undefined); }
			return true;
		} finally { this.deciding = false; }
	}

	protected async _doAccept() { if (!await this.decide('keep') && this.data.changes.length) { throw new Error('A review decision is already running.'); } }
	protected async _doReject() { if (!await this.decide('undo') && this.data.changes.length) { throw new Error('A review decision is already running.'); } }
	protected _createUndoRedoElement(_response: IChatResponseModel) { return undefined; }
	protected async _areOriginalAndModifiedIdentical() { await this.refresh; return !this.data.changes.length; }
	protected _createEditorIntegration(editor: IEditorPane): IModifiedFileEntryEditorIntegration {
		const code = getCodeEditor(editor.getControl());
		if (!code) {
			return { currentIndex: constObservable(-1), reveal() { }, next: () => false, previous: () => false, enableAccessibleDiffView() { }, acceptNearestChange: async () => { await this.decide('keep'); }, rejectNearestChange: async () => { await this.decide('undo'); }, toggleDiff: async () => { }, dispose() { } };
		}
		return this._instantiationService.createInstance(ChatEditingCodeEditorIntegration, this, code, this.diffInfo, true);
	}
	async getDiffInfo(): Promise<IDocumentDiff> { await this.refresh; return this.changeService.getDiffInfo(); }
	async acceptAgentEdits() { throw new Error('Word review projections cannot receive text edits. Edit the Word document instead.'); }
	createSnapshot(_session: URI, _request: string | undefined, _stop: string | undefined): ISnapshotEntry {
		return { resource: this.modifiedURI, languageId: 'markdown', snapshotUri: this.originalURI, original: this.originalModel.getValue(), current: this.modifiedModel.getValue(), state: this.state.get(), telemetryInfo: this.telemetryInfo };
	}
	equalsSnapshot(snapshot: ISnapshotEntry | undefined) { return snapshot?.original === this.originalModel.getValue() && snapshot.current === this.modifiedModel.getValue(); }
	async restoreFromSnapshot(_snapshot: ISnapshotEntry) {
		const data = await this.commands.executeCommand<IOfficeDocumentReview>('aiOffice.word.reviewStatus', this.data.uri);
		assertType(data);
		await this.update(data);
	}
	async resetToInitialContent() { await this.decide('undo'); }
	async resetEditTrackerToInitialContent() { await this.decide('keep'); }
	async computeEditsFromSnapshots() { return []; }
	async save() { await this.commands.executeCommand('aiOffice.word.saveDocument', this.data.uri); }
	async revertToDisk() { await this.restoreFromSnapshot(this.createSnapshot(this.telemetryInfo.sessionResource, undefined, undefined)); }
}

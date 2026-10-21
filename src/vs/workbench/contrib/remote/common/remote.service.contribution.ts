/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

// Remote file labels and extension-host IPC are needed without Remote Explorer or Ports UI.
import { IWorkbenchContribution, IWorkbenchContributionsRegistry, WorkbenchPhase, Extensions as WorkbenchExtensions, registerWorkbenchContribution2 } from '../../../common/contributions.js';
import { Registry } from '../../../../platform/registry/common/platform.js';
import { LifecyclePhase } from '../../../services/lifecycle/common/lifecycle.js';
import { ILabelService, ResourceLabelFormatting } from '../../../../platform/label/common/label.js';
import { OperatingSystem, isWeb, OS } from '../../../../base/common/platform.js';
import { Schemas } from '../../../../base/common/network.js';
import { IRemoteAgentService } from '../../../services/remote/common/remoteAgentService.js';
import { ILoggerService } from '../../../../platform/log/common/log.js';
import { Disposable } from '../../../../base/common/lifecycle.js';
import { IDownloadService } from '../../../../platform/download/common/download.js';
import { DownloadServiceChannel } from '../../../../platform/download/common/downloadIpc.js';
import { RemoteLoggerChannelClient } from '../../../../platform/log/common/logIpc.js';

export class LabelContribution implements IWorkbenchContribution {

	static readonly ID = 'workbench.contrib.remoteLabel';

	constructor(
		@ILabelService private readonly labelService: ILabelService,
		@IRemoteAgentService private readonly remoteAgentService: IRemoteAgentService) {
		this.registerFormatters();
	}

	private registerFormatters(): void {
		this.remoteAgentService.getEnvironment().then(remoteEnvironment => {
			const os = remoteEnvironment?.os || OS;
			const formatting: ResourceLabelFormatting = {
				label: '${path}',
				separator: os === OperatingSystem.Windows ? '\\' : '/',
				tildify: os !== OperatingSystem.Windows,
				normalizeDriveLetter: os === OperatingSystem.Windows,
				workspaceSuffix: isWeb ? undefined : Schemas.vscodeRemote
			};
			this.labelService.registerFormatter({
				scheme: Schemas.vscodeRemote,
				formatting
			});

			if (remoteEnvironment) {
				this.labelService.registerFormatter({
					scheme: Schemas.vscodeUserData,
					formatting
				});
			}
		});
	}
}

class RemoteChannelsContribution extends Disposable implements IWorkbenchContribution {

	constructor(
		@IRemoteAgentService remoteAgentService: IRemoteAgentService,
		@IDownloadService downloadService: IDownloadService,
		@ILoggerService loggerService: ILoggerService,
	) {
		super();
		const connection = remoteAgentService.getConnection();
		if (connection) {
			connection.registerChannel('download', new DownloadServiceChannel(downloadService));
			connection.withChannel('logger', async channel => this._register(new RemoteLoggerChannelClient(loggerService, channel)));
		}
	}
}

registerWorkbenchContribution2(LabelContribution.ID, LabelContribution, WorkbenchPhase.BlockStartup);
Registry.as<IWorkbenchContributionsRegistry>(WorkbenchExtensions.Workbench).registerWorkbenchContribution(RemoteChannelsContribution, LifecyclePhase.Restored);

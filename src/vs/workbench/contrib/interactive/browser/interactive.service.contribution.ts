/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

// Extension-host and chat APIs need these services without the feature UI.
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { IInteractiveDocumentService, InteractiveDocumentService } from './interactiveDocumentService.js';
import { IInteractiveHistoryService, InteractiveHistoryService } from './interactiveHistoryService.js';

registerSingleton(IInteractiveHistoryService, InteractiveHistoryService, InstantiationType.Delayed);
registerSingleton(IInteractiveDocumentService, InteractiveDocumentService, InstantiationType.Delayed);

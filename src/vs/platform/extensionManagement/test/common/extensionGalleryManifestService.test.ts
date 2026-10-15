/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import product from '../../../product/common/product.js';
import { ExtensionGalleryManifestStatus } from '../../common/extensionGalleryManifest.js';
import { ExtensionGalleryManifestService } from '../../common/extensionGalleryManifestService.js';

suite('Extension Gallery Marketplace Availability', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	for (const enabled of [false, true, undefined]) {
		test(`product setting ${enabled} controls gallery availability`, async () => {
			const service = store.add(new ExtensionGalleryManifestService({
				...product,
				_serviceBrand: undefined,
				extensionsMarketplaceEnabled: enabled,
				extensionsGallery: {
					serviceUrl: 'https://example.com/gallery',
					controlUrl: '',
					extensionUrlTemplate: '',
					resourceUrlTemplate: '',
					nlsBaseUrl: '',
				},
			}));
			assert.deepStrictEqual({
				status: service.extensionGalleryManifestStatus,
				hasManifest: (await service.getExtensionGalleryManifest()) !== null,
			}, {
				status: enabled === false ? ExtensionGalleryManifestStatus.Unavailable : ExtensionGalleryManifestStatus.Available,
				hasManifest: enabled !== false,
			});
		});
	}
});

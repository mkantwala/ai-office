/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../base/test/common/utils.js';
import { ConditionalAuthState, conditionalAuthState, resolveSignedOutWindowGate, shouldShowGitHubWorkspaceGroupSignIn, SignedOutWindowGate } from '../../browser/sessionsAuthGate.js';
import { SessionTypeAuthRequirement } from '../../services/sessions/common/session.js';

suite('Sessions - Auth Gate', () => {

	ensureNoDisposablesAreLeakedInTestSuite();

	test('AI Office opens signed out: the gate never forces GitHub sign-in', () => {
		assert.deepStrictEqual({
			featureDisabled: resolveSignedOutWindowGate(false, [SessionTypeAuthRequirement.None]),
			featureDisabledBeforeProviders: resolveSignedOutWindowGate(false, []),
			initializesWithoutGitHub: resolveSignedOutWindowGate(false, [SessionTypeAuthRequirement.GitHub], true),
			initializationBeforeProviders: resolveSignedOutWindowGate(false, [], true),
			providersUnresolved: resolveSignedOutWindowGate(true, []),
			allRequireGitHub: resolveSignedOutWindowGate(true, [SessionTypeAuthRequirement.GitHub, SessionTypeAuthRequirement.GitHub]),
			nativeProvider: resolveSignedOutWindowGate(true, [SessionTypeAuthRequirement.GitHub, SessionTypeAuthRequirement.None]),
			nativeProviderInitializing: resolveSignedOutWindowGate(true, [SessionTypeAuthRequirement.GitHub, SessionTypeAuthRequirement.Unusable]),
		}, {
			featureDisabled: SignedOutWindowGate.Proceed,
			featureDisabledBeforeProviders: SignedOutWindowGate.Unresolved,
			initializesWithoutGitHub: SignedOutWindowGate.Proceed,
			initializationBeforeProviders: SignedOutWindowGate.Unresolved,
			providersUnresolved: SignedOutWindowGate.Unresolved,
			allRequireGitHub: SignedOutWindowGate.Proceed,
			nativeProvider: SignedOutWindowGate.Proceed,
			nativeProviderInitializing: SignedOutWindowGate.Proceed,
		});
	});

	test('the GitHub workspace group never offers sign-in', () => {
		assert.deepStrictEqual([
			shouldShowGitHubWorkspaceGroupSignIn(false, false),
			shouldShowGitHubWorkspaceGroupSignIn(false, true),
			shouldShowGitHubWorkspaceGroupSignIn(true, false),
			shouldShowGitHubWorkspaceGroupSignIn(true, true),
		], [false, false, false, false]);
	});

	test('conditionalAuthState treats an unresolved account as unknown, never signed out', () => {
		// The root-cause distinction: before the account resolves, its snapshot is
		// null for signed-in and signed-out users alike, so `accountResolved: false`
		// must map to Unresolved regardless of the (untrustworthy) signedIn snapshot —
		// otherwise the conditional-auth UI flashes a sign-in modal at a signed-in
		// user during startup.
		const cases = [
			{ accountResolved: false, signedIn: false },
			{ accountResolved: false, signedIn: true },
			{ accountResolved: true, signedIn: false },
			{ accountResolved: true, signedIn: true },
		];

		assert.deepStrictEqual(cases.map(c => conditionalAuthState(c.accountResolved, c.signedIn)), [
			ConditionalAuthState.Unresolved,
			ConditionalAuthState.Unresolved,
			ConditionalAuthState.SignedOut,
			ConditionalAuthState.SignedIn,
		]);
	});
});

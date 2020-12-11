// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('assert');
const nock = require('nock');
const sinon = require('sinon');
const { CloudAdapter, TurnContext } = require('..');
const { jwt, oauth, sinonExt } = require('botbuilder-test-utils');

const {
    AuthenticationConfiguration,
    AuthenticationConstants,
    ClaimsIdentity,
    EmulatorValidation,
    ParameterizedBotFrameworkAuthentication,
    PasswordServiceClientCredentialFactory,
    SkillValidation,
    makeAuthValidator,
} = require('botframework-connector');

const reference = {
    activityId: '1234',
    channelId: 'test',
    serviceUrl: 'https://service.url',
    user: { id: 'user', name: 'User Name' },
    bot: { id: 'bot', name: 'Bot Name' },
    conversation: {
        id: 'convo1',
        properties: {
            foo: 'bar',
        },
    },
};

const incomingMessage = TurnContext.applyConversationReference({ text: 'test', type: 'message' }, reference, true);
const outgoingMessage = TurnContext.applyConversationReference({ text: 'test', type: 'message' }, reference);
// const incomingInvoke = TurnContext.applyConversationReference({ type: 'invoke' }, reference, true);

describe('CloudAdapter', () => {
    jwt.mocha();
    oauth.mocha();

    const sandbox = sinonExt.mocha();

    const makeAdapter = async ({
        appId = 'appId',
        appPassword = 'password',
        callerId = 'caller:id',
        loginUrl = 'login.url',
        oauthScope = 'oauth:scope',
        skipJwtVerify = false,
    } = {}) => {
        const { issuer, metadata, sign, verify } = jwt.stub();

        const credentialFactory = new PasswordServiceClientCredentialFactory(appId, appPassword);
        const authConfig = new AuthenticationConfiguration();

        const authValidator = (expectedClaim) =>
            makeAuthValidator(
                {
                    issuer,
                },
                metadata,
                async (_, identity) => {
                    assert(
                        identity.claims.find((c) => c.type === expectedClaim.type && c.value === expectedClaim.value),
                        'could not locate expected claim'
                    );
                }
            );

        const auth = new ParameterizedBotFrameworkAuthentication(
            credentialFactory,
            authConfig,
            true,
            loginUrl,
            oauthScope,
            callerId,
            authValidator({ type: 'test', value: 'general' }),
            authValidator({ type: 'test', value: 'emulator' }),
            authValidator({ type: 'test', value: 'skill' })
        );

        const adapter = new CloudAdapter(auth);

        return {
            adapter,
            sign,
            verify: () => {
                if (!skipJwtVerify) verify();
                sandbox().verify();
            },
        };
    };

    describe('processActivity', () => {
        it('works', async () => {
            const { adapter, sign, verify } = await makeAdapter();

            const sendActivityNock = nock(reference.serviceUrl)
                .post(`/v3/conversations/${reference.conversation.id}/activities/${reference.activityId}`)
                .reply(200, {});

            const authHeader = `Bearer ${sign({ test: 'general' })}`;

            await adapter.processActivity(authHeader, incomingMessage, async (turnContext) => {
                sinon.assert.match(turnContext, sinon.match.instanceOf(TurnContext));
                await turnContext.sendActivity(outgoingMessage);
            });

            verify();
            sendActivityNock.done();
        });

        it('works for emulator auth', async () => {
            const { adapter, sign, verify } = await makeAdapter();

            const authHeader = `Bearer ${sign({ test: 'emulator' })}`;
            sandbox()
                .mock(EmulatorValidation)
                .expects('isTokenFromEmulator')
                .withArgs(authHeader)
                .once()
                .resolves(true);

            const sendActivityNock = nock(reference.serviceUrl)
                .post(`/v3/conversations/${reference.conversation.id}/activities/${reference.activityId}`)
                .reply(200, {});

            await adapter.processActivity(authHeader, incomingMessage, async (turnContext) => {
                sinon.assert.match(turnContext, sinon.match.instanceOf(TurnContext));
                await turnContext.sendActivity(outgoingMessage);
            });

            verify();
            sendActivityNock.done();
        });

        it('works for skills auth', async () => {
            const { adapter, sign, verify } = await makeAdapter();

            const authHeader = `Bearer ${sign({ test: 'skill' })}`;
            sandbox().mock(SkillValidation).expects('isSkillToken').withArgs(authHeader).once().resolves(true);

            const sendActivityNock = nock(reference.serviceUrl)
                .post(`/v3/conversations/${reference.conversation.id}/activities/${reference.activityId}`)
                .reply(200, {});

            await adapter.processActivity(authHeader, incomingMessage, async (turnContext) => {
                sinon.assert.match(turnContext, sinon.match.instanceOf(TurnContext));
                await turnContext.sendActivity(outgoingMessage);
            });

            verify();
            sendActivityNock.done();
        });
    });

    describe('processProactive', () => {
        it('works', async () => {
            const appId = 'proactiveAppId';
            const oauthScope = 'oauth:scope';

            const { adapter, verify } = await makeAdapter({ appId, oauthScope, skipJwtVerify: true });
            const { match, verify: verifyOauth } = oauth.stub({ clientId: appId, scope: oauthScope });

            const identity = new ClaimsIdentity(
                [
                    {
                        type: AuthenticationConstants.AppIdClaim,
                        value: appId,
                    },
                ],
                true
            );

            const sendActivityNock = match(
                nock(reference.serviceUrl)
                    .post(`/v3/conversations/${reference.conversation.id}/activities/${reference.activityId}`)
                    .reply(200, {})
            );

            await adapter.processProactive(identity, reference, oauthScope, async (turnContext) => {
                sinon.assert.match(turnContext, sinon.match.instanceOf(TurnContext));
                await turnContext.sendActivity(outgoingMessage);
            });

            verify();
            verifyOauth();
            sendActivityNock.done();
        });
    });
});

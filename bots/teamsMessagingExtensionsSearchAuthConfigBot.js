// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');

const {
    ExtendedUserTokenProvider
} = require('botbuilder-core')

const axios = require('axios');
const querystring = require('querystring');
const { SimpleGraphClient } = require('..\\simpleGraphClient.js');
const { polyfills } = require('isomorphic-fetch');

// User Configuration property name
const USER_CONFIGURATION = 'userConfigurationProperty';

class TeamsMessagingExtensionsSearchAuthConfigBot extends TeamsActivityHandler {
    /**
     *
     * @param {UserState} User state to persist configuration settings
     */
    constructor(userState) {
        super();
        // Creates a new user property accessor.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
        this.userConfigurationProperty = userState.createProperty(
            USER_CONFIGURATION
        );
        this.connectionName = process.env.ConnectionName;
        this.userState = userState;
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save state changes
        await this.userState.saveChanges(context);
    }

    // Overloaded function. Receives invoke activities with Activity name of 'composeExtension/queryLink'
    async handleTeamsAppBasedLinkQuery(context, query) {
        const magicCode =
            query.state && Number.isInteger(Number(query.state))
                ? query.state
                : '';
        const tokenResponse = await context.adapter.getUserToken(
            context,
            this.connectionName,
            magicCode
        );

        if (!tokenResponse || !tokenResponse.token) {
            // There is no token, so the user has not signed in yet.

            // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
            const signInLink = await context.adapter.getSignInLink(
                context,
                this.connectionName
            );

            return {
                composeExtension: {
                    type: 'auth',
                    suggestedActions: {
                        actions: [
                            {
                                type: 'openUrl',
                                value: signInLink,
                                title: 'Bot Service OAuth'
                            },
                        ],
                    },
                },
            };
        }
        const graphClient = new SimpleGraphClient(tokenResponse.token);
        const profile = await graphClient.GetMyProfile();
        const userPhoto = await graphClient.GetPhotoAsync(tokenResponse.token);
        const attachment = CardFactory.thumbnailCard(profile.displayName, CardFactory.images([userPhoto]));
        const result = {
            attachmentLayout: 'list',
            type: 'result',
            attachments: [attachment]
        };

        const response = {
            composeExtension: result
        };
        return response;
    }

    async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(
        context,
        query
    ) {
        // The user has requested the Messaging Extension Configuration page settings url.
        const userSettings = await this.userConfigurationProperty.get(
            context,
            ''
        );
        const escapedSettings = userSettings
            ? querystring.escape(userSettings)
            : '';

        return {
            composeExtension: {
                type: 'config',
                suggestedActions: {
                    actions: [
                        {
                            type: ActionTypes.OpenUrl,
                            value: `${process.env.SiteUrl}/public/searchSettings.html?settings=${escapedSettings}`
                        },
                    ],
                },
            },
        };
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/setting
    async handleTeamsMessagingExtensionConfigurationSetting(context, settings) {
        // When the user submits the settings page, this event is fired.
        if (settings.state != null) {
            await this.userConfigurationProperty.set(context, settings.state);
        }
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/query'.
    async handleTeamsMessagingExtensionQuery(context, query) {
        const searchQuery = query.parameters[0].value;
        const attachments = [];
        const userSettings = await this.userConfigurationProperty.get(
            context,
            ''
        );


        if (userSettings.includes('profile')) {
            // When the Bot Service Auth flow completes, the query.State will contain a magic code used for verification.
            const magicCode =
                query.state && Number.isInteger(Number(query.state))
                    ? query.state
                    : '';
            const tokenResponse = await context.adapter.getUserToken(
                context,
                this.connectionName,
                magicCode
            );

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.

                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                const signInLink = await context.adapter.getSignInLink(
                    context,
                    this.connectionName
                );

                return {
                    composeExtension: {
                        type: 'silentAuth',
                        suggestedActions: {
                            actions: [
                                {
                                    type: 'openUrl',
                                    value: signInLink,
                                    title: 'Bot Service OAuth'
                                },
                            ],
                        },
                    },
                };
            }

            // The user is signed in, so use the token to create a Graph Clilent and show profile
            console.log(tokenResponse.token);
            const graphClient = new SimpleGraphClient(tokenResponse.token);
            const profile = await graphClient.GetMyProfile();
            const userPhoto = await graphClient.GetPhotoAsync(tokenResponse.token);
            const thumbnailCard = CardFactory.thumbnailCard(profile.displayName, CardFactory.images([userPhoto]));
            attachments.push(thumbnailCard);
        } else if (userSettings.includes('customwebserviceanon')) {
            const response = await axios.get(
                `https://msteamssamples-httpsample.azurewebsites.net/api/HttpSample?${querystring.stringify({
                    name: searchQuery
                })}`
            );

            response.data.objects.forEach((obj) => {
                const heroCard = CardFactory.heroCard(obj.item.name);
                const preview = CardFactory.heroCard(obj.item.name);
                preview.content.tap = {
                    type: 'invoke',
                    value: { description: obj.item.description }
                };
                attachments.push({ ...heroCard, preview });
            });
        } else if (userSettings.includes('customwebserviceauth')) {
            const accesstokenEndpoint = `https://login.microsoftonline.com/${process.env.TenantId}/oauth2/token`
            const formData = new URLSearchParams();
            formData.append('grant_type', 'client_credentials');
            formData.append('client_id', process.env.MicrosoftAppId);
            formData.append('client_secret', process.env.MicrosoftAppPassword);
            formData.append('resource', process.env.MicrosoftAppId);

            const headerConfig = {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
            const accesstokenresponse = await axios.post(accesstokenEndpoint, formData, headerConfig);
            console.log("Azure Function Access Token: " + accesstokenresponse.data.access_token)


            const customAPIEndpoint = `https://msteamssamples-httpsampleauth.azurewebsites.net/api/HttpSampleAuth?${querystring.stringify({ name: searchQuery })}`;
            const response = await axios.get(customAPIEndpoint, {
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "bearer " + accesstokenresponse.data.access_token
                }
            });

            response.data.objects.forEach((obj) => {
                const heroCard = CardFactory.heroCard(obj.item.name);
                const preview = CardFactory.heroCard(obj.item.name);
                preview.content.tap = {
                    type: 'invoke',
                    value: { description: obj.item.description }
                };
                attachments.push({ ...heroCard, preview });
            });
        } else {
            const response = await axios.get(
                `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
                    text: searchQuery,
                    size: 8
                })}`
            );

            response.data.objects.forEach((obj) => {
                const heroCard = CardFactory.heroCard(obj.package.name);
                const preview = CardFactory.heroCard(obj.package.name);
                preview.content.tap = {
                    type: 'invoke',
                    value: { description: obj.package.description }
                };
                attachments.push({ ...heroCard, preview });
            });
        }

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: attachments
            },
        };
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/selectItem'.
    async handleTeamsMessagingExtensionSelectItem(context, obj) {
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [CardFactory.thumbnailCard(obj.description)]
            },
        };
    }

    // Overloaded function. Receives invoke activities with the name 'composeExtension/fetchTask'
    async handleTeamsMessagingExtensionFetchTask(context, action) {
        if (action.commandId === 'SHOWPROFILE') {
            const magicCode =
                action.state && Number.isInteger(Number(action.state))
                    ? action.state
                    : '';
            const tokenResponse = await context.adapter.getUserToken(
                context,
                this.connectionName,
                magicCode
            );

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

                const signInLink = await context.adapter.getSignInLink(
                    context,
                    this.connectionName
                );

                return {
                    composeExtension: {
                        type: 'silentAuth',
                        suggestedActions: {
                            actions: [
                                {
                                    type: 'openUrl',
                                    value: signInLink,
                                    title: 'Bot Service OAuth'
                                },
                            ],
                        },
                    },
                };
            }
            const graphClient = new SimpleGraphClient(tokenResponse.token);
            const profile = await graphClient.GetMyProfile();
            const userPhoto = await graphClient.GetPhotoAsync(tokenResponse.token);
            const profileCard = CardFactory.adaptiveCard({
                version: '1.0.0',
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'TextBlock',
                        text: 'Hello: ' + profile.displayName,
                    },
                    {
                        type: 'Image',
                        url: userPhoto,
                    },
                ],
            });
            return {
                task: {
                    type: 'continue',
                    value: {
                        card: profileCard,
                        heigth: 250,
                        width: 400,
                        title: 'Show Profile Card'
                    },
                },
            };
        }
        if (action.commandId === 'SignOutCommand') {
            const adapter = context.adapter;
            await adapter.signOutUser(context, this.connectionName);

            const card = CardFactory.adaptiveCard({
                version: '1.0.0',
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'TextBlock',
                        text: 'You have been signed out.'
                    },
                ],
                actions: [
                    {
                        type: 'Action.Submit',
                        title: 'Close',
                        data: {
                            key: 'close'
                        },
                    },
                ],
            });

            return {
                task: {
                    type: 'continue',
                    value: {
                        card: card,
                        heigth: 200,
                        width: 400,
                        title: 'Adaptive Card: Inputs'
                    },
                },
            };
        }
        if (action.commandId === 'shareMessage') {
            const magicCode =
                action.state && Number.isInteger(Number(action.state))
                    ? action.state
                    : '';
            const tokenResponse = await context.adapter.getUserToken(
                context,
                this.connectionName,
                magicCode
            );

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

                const signInLink = await context.adapter.getSignInLink(
                    context,
                    this.connectionName
                );

                return {
                    composeExtension: {
                        type: 'silentAuth',
                        suggestedActions: {
                            actions: [
                                {
                                    type: 'openUrl',
                                    value: signInLink,
                                    title: 'Bot Service OAuth'
                                },
                            ],
                        },
                    },
                };
            }
            const graphClient = new SimpleGraphClient(tokenResponse.token);
            const CopyMessageResponse = await graphClient.SendMessage(tokenResponse.token, action);
            console.log(CopyMessageResponse);
        }
        return null;
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        switch (action.commandId) {
            case "shareMessage":
                return {}; //shareMessageCommand(context, action);
            default:
                return {};
        }
    }

    async onInvokeActivity(context) {
        console.log('onInvoke, ' + context.activity.name);
        const valueObj = context.activity.value;
        if (valueObj.authentication) {
            const authObj = valueObj.authentication;
            if (authObj.token) {
                // If the token is NOT exchangeable, then do NOT deduplicate requests.
                if (await this.tokenIsExchangeable(context)) {
                    return await super.onInvokeActivity(context);
                }
                else {
                    const response = {
                        status: 412
                    };
                    return response;
                }
            }
        }
        return await super.onInvokeActivity(context);
    }

    async tokenIsExchangeable(context) {
        let tokenExchangeResponse = null;
        try {
            const valueObj = context.activity.value;
            const tokenExchangeRequest = valueObj.authentication;
            console.log("tokenExchangeRequest.token: " + tokenExchangeRequest.token);

            tokenExchangeResponse = await context.adapter.exchangeToken(context,
                process.env.connectionName,
                context.activity.from.id,
                { token: tokenExchangeRequest.token });
            console.log('tokenExchangeResponse: ' + JSON.stringify(tokenExchangeResponse));
        } catch (err) {
            console.log('tokenExchange error: ' + err);
            // Ignore Exceptions
            // If token exchange failed for any reason, tokenExchangeResponse above stays null , and hence we send back a failure invoke response to the caller.
        }
        if (!tokenExchangeResponse || !tokenExchangeResponse.token) {
            return false;
        }

        console.log('Exchanged token: ' + tokenExchangeResponse.token);
        return true;
    }

}

function shareMessageCommand(context, action) {
    const accesstokenEndpoint1 = `https://login.microsoftonline.com/${process.env.TenantId}/oauth2/token`
    const formData1 = new URLSearchParams();
    formData1.append('grant_type', 'client_credentials');
    formData1.append('client_id', process.env.MicrosoftAppId);
    formData1.append('client_secret', process.env.MicrosoftAppPassword);
    formData1.append('resource', process.env.MicrosoftAppId);

    const headerConfig1 = {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }
    const accesstokenresponse1 = axios.post(accesstokenEndpoint1, formData1, headerConfig1);
    console.log("Azure Function Access Token: " + accesstokenresponse1.data.access_token);

    const heroCard = CardFactory.heroCard(
        "originally sent this message:",
        action.messagePayload.body.content,
        null
    );

    const customAPIEndpoint = "https://graph.microsoft.com/v1.0/teams/9cfc4d00-a0da-403c-823f-3e384fb6e9d0/channels/19:joA8BNKHbJ8s9NM4qtrcdAenhoXK_xByUdlJOo1EBRs1@thread.tacv2/messages";
    const chatmessage = {
        "subject": null,
        "body": {
            "contentType": "html",
            "content": "<attachment id=\"74d20c7f34aa4a7fb74e2b30004247c5\"></attachment>"
        },
        "attachments": [
            {
                "id": "74d20c7f34aa4a7fb74e2b30004247c5",
                "contentType": heroCard.contentType,
                "content": heroCard.content,
                "name": null,
                "thumbnailUrl": null
            }
        ]

    };

    const response = axios.post(customAPIEndpoint, chatmessage, {
        headers: {
            "Content-Type": "application/json",
            "Authorization": "bearer " + accesstokenresponse1.data.access_token
        }
    });

    return {};
}

module.exports.TeamsMessagingExtensionsSearchAuthConfigBot = TeamsMessagingExtensionsSearchAuthConfigBot;

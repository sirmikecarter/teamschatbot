// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {TurnContext, MessageFactory, TeamsInfo, TeamsActivityHandler, CardFactory, ActionTypes} = require('botbuilder');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');
const axios = require('axios');
const querystring = require('querystring');
const TextEncoder = require('util').TextEncoder;

const { DialogHelper } = require('./helpers/dialogHelper');

const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;

class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();

        const luisApplication = {
            applicationId: process.env.LuisAppId,
            azureRegion: process.env.LuisAPIHostName,
            // CAUTION: Authoring key is used in this example as it is appropriate for prototyping.
            // When implimenting for deployment/production, assign and use a subscription key instead of an authoring key.
            endpointKey: process.env.LuisAPIKey
        };

        const luisPredictionOptions = {
            spellCheck: true,
            bingSpellCheckSubscriptionKey: process.env.BingSpellCheck

        };

        this.qnaRecognizer = new QnAMaker({
            knowledgeBaseId: process.env.QnAKbId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAHostname
        });

        this.luisRecognizer = new LuisRecognizer(luisApplication, luisPredictionOptions);
        this.dialogHelper = new DialogHelper();

        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            switch (context.activity.text.trim()) {
            case 'MentionMe':
                await this.mentionActivityAsync(context);
                break;
            case 'UpdateCardAction':
                await this.updateCardActivityAsync(context);
                break;
            case 'Delete':
                await this.deleteCardActivityAsync(context);
                break;
            case 'MessageAllMembers':
                await this.messageAllMembersAsync(context);
                break;
            case 'Select a Term':
                await context.sendActivity(`This is under Construction`);
                break;
            case 'See All Terms':
                await context.sendActivity(`This is under Construction`);
                break;
            case 'Glossary Search':
                await context.sendActivity(`This is under Construction`);
                break;
            default:

            const dispatchResults = await this.luisRecognizer.recognize(context);
            const dispatchTopIntent = LuisRecognizer.topIntent(dispatchResults);

            switch (dispatchTopIntent) {
              case 'General':
                  const qnaResult = await this.qnaRecognizer.generateAnswer(dispatchResults.text, QNA_TOP_N, QNA_CONFIDENCE_THRESHOLD);
                  if (!qnaResult || qnaResult.length === 0 || !qnaResult[0].answer){
                    //await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }else{
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }

                  break;
              case 'Glossary':
                  const glossaryTerm = dispatchResults.entities.Term[0];
                  await context.sendActivity(glossaryTerm);
                  break;

            }

            //await context.sendActivity(dispatchTopIntent);


                const value = { count: 0 };
                const card = CardFactory.heroCard(
                    'What would you like to do?',
                    null,
                    [
                        {
                            type: ActionTypes.MessageBack,
                            title: 'Select a Term',
                            value: value,
                            text: 'Select a Term'
                        },
                        {
                            type: ActionTypes.MessageBack,
                            title: 'See All Terms',
                            value: null,
                            text: 'See All Terms'
                        },
                        {
                            type: ActionTypes.MessageBack,
                            title: 'Glossary Search',
                            value: null,
                            text: 'Glossary Search'
                        }]);
                await context.sendActivity({ attachments: [card] });
                break;
            }
            await next();
        });

        this.onMembersAddedActivity(async (context, next) => {
            context.activity.membersAdded.forEach(async (teamMember) => {
                if (teamMember.id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to the team ${ teamMember.givenName } ${ teamMember.surname }`);
                }
            });
            await next();
        });
    }

    async mentionActivityAsync(context) {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }

    async updateCardActivityAsync(context) {
        const data = context.activity.value;
        data.count += 1;

        const card = CardFactory.heroCard(
            'Welcome Card',
            `Updated count - ${ data.count }`,
            null,
            [
                {
                    type: ActionTypes.MessageBack,
                    title: 'Update Card',
                    value: data,
                    text: 'UpdateCardAction'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Message all members',
                    value: null,
                    text: 'MessageAllMembers'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Delete card',
                    value: null,
                    text: 'Delete'
                }
            ]);

        card.id = context.activity.replyToId;
        await context.updateActivity({ attachments: [card], id: context.activity.replyToId, type: 'message' });
    }

    async deleteCardActivityAsync(context) {
        await context.deleteActivity(context.activity.replyToId);
    }

    async messageAllMembersAsync(context) {
        const members = await TeamsInfo.getMembers(context);

        members.forEach(async (teamMember) => {
            const message = MessageFactory.text(`Hello ${ teamMember.givenName } ${ teamMember.surname }. I'm a Teams conversation bot.`);

            var ref = TurnContext.getConversationReference(context.activity);
            ref.user = teamMember;

            await context.adapter.createConversation(ref,
                async (t1) => {
                    const ref2 = TurnContext.getConversationReference(t1.activity);
                    await t1.adapter.continueConversation(ref2, async (t2) => {
                        await t2.sendActivity(message);
                    });
                });
        });

        await context.sendActivity(MessageFactory.text('All messages have been sent.'));
    }

}

module.exports.TeamsConversationBot = TeamsConversationBot;

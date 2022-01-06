// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsInfo , TeamsActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const cards = require('./adaptiveCards/cards')
const aad = require('./aad_auth');

class VideoPlayerBot extends TeamsActivityHandler {
    constructor(){
        super();
        
        this.onEvent(async (context, next) => {
            console.log(context);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        })

        this.onTurn(async (context, next) => {
            console.log("Activity received of type: ", context.activity.type);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        })

        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            let msg = '';
            if ( 'text' in context.activity) { msg = context.activity.text.trim().toLowerCase() }
            else if ( 'value' in context.activity) { msg = context.activity.value.type }

            const aadObjectId = context.activity.from.aadObjectId;
            const TeamsExternalAppID = process.env.TeamsExternalAppID;
            const accessToken = await aad.getAccessToken();
            const appInfo = await aad.getTeamsAppID(accessToken, aadObjectId, TeamsExternalAppID);
            if (appInfo.length ===0) { msg = 'error-appid' };

            switch (msg) {
                case 'hello':
                    // Send choice card
                    const choiceCard = cards.getStaticCard('choiceCard', appInfo.id);
                    const choiceCardAC = CardFactory.adaptiveCard(choiceCard);
                    await context.sendActivity(MessageFactory.attachment(choiceCardAC));   
                    break;

                case 'appid': 
                    // Send application access token      
                    const appInfoAC = CardFactory.adaptiveCard({
                        "type": "AdaptiveCard",
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.0",
                        "body": [
                            {
                                "type": "FactSet",
                                "facts": [
                                    { "title": "TeamsAppID",        "value": appInfo.id                 },
                                    { "title": "externalId",        "value": appInfo.externalId         },
                                    { "title": "displayName",       "value": appInfo.displayName        },
                                    { "title": "distributionMethod","value": appInfo.distributionMethod }
                                ]
                            }
                        ]
                    });
                    await context.sendActivity(MessageFactory.attachment(appInfoAC));  
                    break;

                case 'demoCard':
                    // Send demo card
                    const demoCard = cards.getStaticCard('demoCard', appInfo.id);
                    const demoCardAC = CardFactory.adaptiveCard(demoCard);
                    await context.sendActivity(MessageFactory.attachment(demoCardAC));   
                    break;

                case 'inputCard':
                    // Send custom card to get video inputs
                    const inputCard = cards.getStaticCard('inputCard', appInfo.id);
                    const inputCardAC = CardFactory.adaptiveCard(inputCard);
                    await context.sendActivity(MessageFactory.attachment(inputCardAC));   
                    break;

                case 'videoInputs':
                    // Send the custom card to play the video
                    const videoCard = cards.generateVideoCard(context.activity.value.videoName, context.activity.value.videoURL, context.activity.value.websiteURL, appInfo.id);
                    const videoCardAC = CardFactory.adaptiveCard(videoCard);
                    await context.sendActivity(MessageFactory.attachment(videoCardAC)); 
                    break;
                
                case 'error-appid':
                    // Send error message - "Could not get internal app ID form MS Graph API"    
                    const errorAC = CardFactory.adaptiveCard({
                        "type": "AdaptiveCard",
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.0",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "Error - Could not get internal app ID form MS Graph API",
                                "wrap": true,
                                "id": "ErrMsgTitle",
                                "color": "Attention",
                                "weight": "Bolder"
                            },
                            {
                                "type": "TextBlock",
                                "text": "- Check that the **app permissions** set to **TeamsAppInstallation.ReadForUser.All** in Azure AD \n- Check that **TeamsExternalAppID** is correctly set in **.env file**",
                                "wrap": true,
                                "id": "ErrMsgText"
                            }
                        ]
                    });
                    await context.sendActivity(MessageFactory.attachment(errorAC));  
                    break;

                default : // Echo                    
                    const replyText = `Echo: ${ msg }`;
                    await context.sendActivity(MessageFactory.text(replyText, replyText)); 
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            const instructionText = 'Please choose the type of card you want to test';
            const choiceCard = cards.getStaticCard('choiceCard');
            const choiceCardAC = CardFactory.adaptiveCard(choiceCard);
   
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                    await context.sendActivity(MessageFactory.text(instructionText, instructionText));
                    await context.sendActivity(MessageFactory.attachment(choiceCardAC));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });


    }

}

module.exports.VideoPlayerBot = VideoPlayerBot;
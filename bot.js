// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// bot.js is your main bot dialog entry point for handling activity types

// Import required Bot Builder
const { ActionTypes, ActivityTypes, CardFactory } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { DialogSet, WaterfallDialog } = require('botbuilder-dialogs');
const { OAuthHelpers, LOGIN_PROMPT } = require('./oauth-helpers');

const CONNECTION_SETTING_NAME = '<MS Graph API Connection Name>';

/**
 * Demonstrates the following concepts:
 *  Displaying a Welcome Card, using Adaptive Card technology
 *  Use LUIS to model Greetings, Help, and Cancel interactions
 *  Use a Waterfall dialog to model multi-turn conversation flow
 *  Use custom prompts to validate user input
 *  Store conversation and user state
 *  Handle conversation interruptions
 */
let luisResult = null;
class BasicBot {
    /**
     * Constructs the three pieces necessary for this bot to operate:
     * 1. StatePropertyAccessor for conversation state
     * 2. StatePropertyAccess for user state
     * 3. LUIS client
     * 4. DialogSet to handle our GreetingDialog
     *
     * @param {ConversationState} conversationState property accessor
     
     */

    
    constructor(conversationState, application, luisPredictionOptions, includeApiResults) {       
       
        this.luisRecognizer = new LuisRecognizer(application,luisPredictionOptions, true);
        this.conversationState = conversationState;

        // DialogState property accessor. Used to keep persist DialogState when using DialogSet.
        this.dialogState = conversationState.createProperty('dialogState');
        this.commandState = conversationState.createProperty('commandState');
       

        // Instructions for the user with information about commands that this bot may handle.
        this.helpMessage = `You can type "send <recipient_email>" to send an email, "recent" to view recent unread mail,` +
            ` "me" to see information about your, or "help" to view the commands` +
            ` again. For others LUIS displays intent with score.`;

        // Create a DialogSet that contains the OAuthPrompt.
        this.dialogs = new DialogSet(this.dialogState);

        // Add an OAuthPrompt with the connection name as specified on the Bot's settings blade in Azure.
        this.dialogs.add(OAuthHelpers.prompt(CONNECTION_SETTING_NAME));

        this._graphDialogId = 'graphDialog';

        // Logs in the user and calls proceeding dialogs, if login is successful.
        this.dialogs.add(new WaterfallDialog(this._graphDialogId, [
            this.promptStep.bind(this),
            this.processStep.bind(this)
]));

    }

// constructor(conversationState) {
//     this.conversationState = conversationState;

//     // DialogState property accessor. Used to keep persist DialogState when using DialogSet.
//     this.dialogState = conversationState.createProperty('dialogState');
//     this.commandState = conversationState.createProperty('commandState');

//     // Instructions for the user with information about commands that this bot may handle.
//     this.helpMessage = `You can type "send <recipient_email>" to send an email, "recent" to view recent unread mail,` +
//         ` "me" to see information about your, or "help" to view the commands` +
//         ` again. Any other text will display your token.`;

//     // Create a DialogSet that contains the OAuthPrompt.
//     this.dialogs = new DialogSet(this.dialogState);

//     // Add an OAuthPrompt with the connection name as specified on the Bot's settings blade in Azure.
//     this.dialogs.add(OAuthHelpers.prompt(CONNECTION_SETTING_NAME));

//     this._graphDialogId = 'graphDialog';

//     // Logs in the user and calls proceeding dialogs, if login is successful.
//     this.dialogs.add(new WaterfallDialog(this._graphDialogId, [
//         this.promptStep.bind(this),
//         this.processStep.bind(this)
//     ]));
// };

   

    /**
     * Driver code that does one of the following:
     * 1. Display a welcome card upon receiving ConversationUpdate activity
     * 2. Use LUIS to recognize intents for incoming user message
     * 3. Start a greeting dialog
     * 4. Optionally handle Cancel or Help interruptions
     *
     * @param {Context} context turn context from the adapter
     */
    async onTurn(turnContext) {
        //console.log(turnContext);
        const dc = await this.dialogs.createContext(turnContext);
        const results = await this.luisRecognizer.recognize(turnContext);
            
        switch (turnContext._activity.type) {

            case ActivityTypes.Message:

            this.luisResult = results;
            
            //if (topIntent.intent !== 'None' && topIntent !== "Greeting") 
               // await turnContext.sendActivity(`LUIS Top Scoring Intent: ${ topIntent.intent }, Score: ${ topIntent.score }`);
            
                await this.processInput(dc);
                
                    
               
                break;
            case ActivityTypes.Event:
            case ActivityTypes.Invoke:
                if (turnContext._activity.type === ActivityTypes.Invoke && turnContext._activity.channelId !== 'msteams') {
                    throw new Error('The Invoke type is only valid on the MS Teams channel.');
                };
                await dc.continueDialog();
                if (!turnContext.responded) {
                    await dc.beginDialog(this._graphDialogId);
                };
                break;
                case ActivityTypes.ConversationUpdate:
            await this.sendWelcomeMessage(turnContext);
            break;
        default:
            await turnContext.sendActivity(`[${ turnContext._activity.type }]-type activity detected.`);
        }
        
        //await this.luisState.set(results);
        await this.conversationState.saveChanges(turnContext);
    }
    
    async sendWelcomeMessage(turnContext) {
        const activity = turnContext.activity;
        //console.log(activity);
        if (activity && activity.membersAdded) {
            const heroCard = CardFactory.heroCard(
                'Welcome to LUIS with MSGraph API Authentication BOT!',
                CardFactory.images(['https://botframeworksamples.blob.core.windows.net/samples/aadlogo.png']),
                CardFactory.actions([
                    {
                        type: ActionTypes.ImBack,
                        title: 'Log me in',
                        value: 'login'
                    },
                    {
                        type: ActionTypes.ImBack,
                        title: 'Me',
                        value: 'me'
                    },
                    {
                        type: ActionTypes.ImBack,
                        title: 'Recent',
                        value: 'recent'
                    },
                    {
                        type: ActionTypes.ImBack,
                        title: 'View Token',
                        value: 'viewToken'
                    },
                    {
                        type: ActionTypes.ImBack,
                        title: 'Help',
                        value: 'help'
                    },
                    {
                        type: ActionTypes.ImBack,
                        title: 'Signout',
                        value: 'signout'
                    }
                ])
            );

            for (const idx in activity.membersAdded) {
                if (activity.membersAdded[idx].id !== activity.recipient.id) {
                    await turnContext.sendActivity({ attachments: [heroCard] });
                }
            }
        }
}

    async processInput(dc, luisResult) {
        //console.log(dc);
        switch (dc.context.activity.text.toLowerCase()) {
        case 'signout':
        case 'logout':
        case 'signoff':
        case 'logoff':
            // The bot adapter encapsulates the authentication processes and sends
            // activities to from the Bot Connector Service.
            const botAdapter = dc.context.adapter;
            await botAdapter.signOutUser(dc.context, CONNECTION_SETTING_NAME);
            // Let the user know they are signed out.
            await dc.context.sendActivity('You are now signed out.');
            break;
        case 'help':
            await dc.context.sendActivity(this.helpMessage);
            break;
        default:
            // The user has input a command that has not been handled yet,
            // begin the waterfall dialog to handle the input.
            await dc.continueDialog();
            if (!dc.context.responded) {
                await dc.beginDialog(this._graphDialogId);
            }
            //console.log(luisResult);
        }
};

async promptStep(step) {
    //console.log(step.context.activity);
        const activity = step.context.activity;

        if (activity.type === ActivityTypes.Message && !(/\d{6}/).test(activity.text)) {
            await this.commandState.set(step.context, activity.text);
            await this.conversationState.saveChanges(step.context);
        }
        return await step.beginDialog(LOGIN_PROMPT);
}

async processStep(step) {
    //console.log(step);
        // We do not need to store the token in the bot. When we need the token we can
        // send another prompt. If the token is valid the user will not need to log back in.
        // The token will be available in the Result property of the task.
        const tokenResponse = step.result;
        //console.log(step.context);

        // If the user is authenticated the bot can use the token to make API calls.
        if (tokenResponse !== undefined) {
            let parts = await this.commandState.get(step.context);
            if (!parts) {
                parts = step.context.activity.text;
            }
            const command = parts.split(' ')[0].toLowerCase();
            console.log(command);
            if(command === 'login' || command === 'signin'){
                await step.context.sendActivity(`You have already loggedin!`);
            }
            else if (command === 'me') {
                await OAuthHelpers.listMe(step.context, tokenResponse);
            } else if (command === 'send') {
                await OAuthHelpers.sendMail(step.context, tokenResponse, parts.split(' ')[1].toLowerCase());
            } else if (command === 'recent') {
                await OAuthHelpers.listRecentMail(step.context, tokenResponse);
            } else if(command.toLowerCase() === 'viewtoken'){
                await step.context.sendActivity(`Your token is: ${ tokenResponse.token }`);
            }else{

                console.log(this.luisResult);

                const topIntent = this.luisResult.luisResult.topScoringIntent;
                if(topIntent !== 'None'){
                    await step.context.sendActivity(`LUIS Top Scoring Intent: ${ topIntent.intent }, Score: ${ topIntent.score }`);
                }else{
                   
                await step.context.sendActivity(`Please try something else!`);
                // If the top scoring intent was "None" tell the user no valid intents were found and provide help.
                    // await step.context.sendActivity(`No LUIS intents were found.
                    //                                 \nThis sample is about identifying two user intents:
                    //                                 \n - 'Calendar.Add'
                    //                                 \n - 'Calendar.Find'
                    //                                 \nTry typing 'Add Event' or 'Show me tomorrow'.`);
            
                }
            }
        } else {
            // Ask the user to try logging in later as they are not logged in.
            await step.context.sendActivity(`We couldn't log you in. Please try again later.`);
        }
        return await step.endDialog();
};

};

exports.BasicBot = BasicBot;

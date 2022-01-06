var urlParser = require('url');

const getStaticCard = (cardName, TeamsInternalAppID) => {
//    console.log(TeamsInternalAppID)
    switch(cardName) {
        case 'choiceCard':
            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.4",
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "Which type of video card do you want to try?"
                    },
                    {
                        "type": "Image",
                        "url": `${process.env.baseURL}/media/demo-video.gif`,
                        "horizontalAlignment": "Center",
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Demo Card",
                        "data": {
                            "type": "demoCard"
                        }
                    },
                    {
                        "type": "Action.Submit",
                        "title": "Custom Card",
                        "data": {
                            "type": "inputCard"
                        }
                    }
                ]
            }     
        case 'demoCard':
            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.3",
                "type": "AdaptiveCard",
                "body": [
                    {
                        "size": "ExtraLarge",
                        "text": "Learn Together",
                        "weight": "Bolder",
                        "wrap": true,
                        "type": "TextBlock"
                    },
                    {
                        "type": "TextBlock",
                        "text": "Developing apps for Microsoft Teams",
                        "wrap": true,
                        "spacing": "None"
                    },
                    {
                        "text": "Click on the image to watch the video in Teams",
                        "weight": "Bolder",
                        "type": "TextBlock",
                        "isSubtle": true
                    },
                    {
                        "type": "Image",
                        "url": `${process.env.baseURL}/media/learn-TV.png`,
                        "horizontalAlignment": "Center",
                        "selectAction": {
                            "type": "Action.OpenUrl",
                            "url": `https://teams.microsoft.com/l/stage/${TeamsInternalAppID}/0?context=${urlEncoder('Learn Together - Developing apps for Microsoft Teams', 'https://www.youtube.com/embed/xxkCJKpU3vA', 'https://www.youtube.com/watch?v=xxkCJKpU3vA')}`                        }
                    },
                    {
                        "text": "Check out the stream on learn TV to ask your questions live: [https://aka.ms/learntv](https://aka.ms/learntv)\n\nCurrently, there are 115+ million Teams daily active users. That is millions of new users that developers can reach when they develop apps for Teams.\n\nJoin us for a free two-hour livestream event for developers by developers. Let's talk app dev for Microsoft Teams.\n\nWhat you will learn:\n\nMillions of new Microsoft Teams users are looking to you, the developers, to create engaging and unique application experiences on Teams. In this two-hour livestream on Learn TV, you’ll learn:\n\n- Why you should consider building apps for Teams\n- How to get started building apps for Teams in VS Code\n- Where you can integrate your apps in the Teams user experience\n\nQuickly get started learning how to build these apps and stick around for some fun trivia and prizes.",
                        "wrap": true,
                        "type": "TextBlock"
                    },
                    {
                        "size": "Small",
                        "text": "Microsoft Developer",
                        "weight": "Lighter",
                        "wrap": true,
                        "type": "TextBlock"
                    }
                ],
                "actions": [
                    {
                        "url": "https://www.youtube.com/watch?v=xxkCJKpU3vA",
                        "title": "Watch the video on YouTube",
                        "type": "Action.OpenUrl"
                    }
                ]
            }
        case 'demoCardtest':
            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "ActionSet",
                        "actions": [
                            {
                                "type": "Action.Submit",
                                "title": "View",
                                "data": {
                                    "msteams": {
                                        "type": "invoke",
                                        "value": {
                                            "type": "tab/tabInfoAction",
                                            "tabInfo": {
                                                "websiteURL": "https://www.youtube.com/embed/f71Fv6t4fl4",
                                                "websiteUrl": "https://www.youtube.com/watch?v=f71Fv6t4fl4",
                                                "name": "Test",
                                                "entityId": "stageView"
                                            }
                                        }
                                    }
                                }
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.3"
            }            
        case 'inputCard':   
            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.4",
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "Please provide content information"
                    },
                    {
                        "type": "Input.Text",
                        "id": "videoName",
                        "isRequired": true,
                        "label": "Video Name",
                        "placeholder": "The name of your video",
                        "errorMessage": "Please make sure you provided a name for your video"
                    },
                    {
                        "type": "Input.Text",
                        "placeholder": "The location of your video in https:// format",
                        "id": "videoURL",
                        "isRequired": true,
                        "label": "Video URL",
                        "errorMessage": "Please make sure you provided a URL for your video"
                    },
                    {
                        "type": "Input.Text",
                        "placeholder": "The URL where the video is located in https:// format",
                        "id": "websiteURL",
                        "label": "Website URL",
                        "isRequired": true,
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Validate",
                        "data": {
                            "type": "videoInputs"
                        }
                    }
                ]
            }
        default :
            return {};
    }
} 

const urlEncoder = (videoName, videoURL, websiteURL) => {
    const urlObject = urlParser.parse(videoURL, false);
    const isSPO = urlObject.host.toString().includes('.sharepoint.com') ? true : false;
    if (isSPO) {
        const TeamsLogon  = '/_layouts/15/teamslogon.aspx?spfx=true&dest=';
        const videoURLSPO = `https://${urlObject.hostname}${TeamsLogon}${urlObject.path}`;
        return encodeURIComponent(`{"contentUrl":"${videoURLSPO}","websiteUrl":"${websiteURL}","name":"${videoName}"}`);
    }
    else {
        return encodeURIComponent(`{"contentUrl":"${videoURL}","websiteUrl":"${websiteURL}","name":"${videoName}"}`);
    }    
}

const generateVideoCard = (videoName, videoURL, websiteURL, TeamsInternalAppID) => ({
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.4",
    "type": "AdaptiveCard",
    "body": [
        {
            "size": "ExtraLarge",
            "text": `${videoName}`,
            "weight": "Bolder",
            "wrap": true,
            "type": "TextBlock"
        },
        {
            "text": `Click on -Play- button to start the video`,
            "weight": "Bolder",
            "type": "TextBlock",
            "isSubtle": true
        },
        {
            "type": "Container",
            "items": [
                {
                    "type": "Image",
                    "url": `${process.env.baseURL}/media/stream-play.png`,
                    "horizontalAlignment": "Center",
                    "selectAction": {
                        "type": "Action.OpenUrl",
                        "url": `https://teams.microsoft.com/l/stage/${TeamsInternalAppID}/0?context=${urlEncoder(videoName, videoURL, websiteURL)}`
                    },
                    "height": "80px"
                }
            ],
            "minHeight": "120px",
            "verticalContentAlignment": "Center"
        }
    ],
    "actions": [
        {
            "url": `${websiteURL}`,
            "title": "Click here to access to the video content",
            "type": "Action.OpenUrl"
        }
    ]
})

module.exports = { getStaticCard, generateVideoCard }
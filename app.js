const TelegramBot = require('node-telegram-bot-api')
// replace with your telegram bot token
const token = 'YOUR_TELEGRAM_BOT_TOKEN'
const bot = new TelegramBot(token, {polling: true})

const axios = require('axios')
const qs = require('qs') // for data serialization



// link used to get access token
// replace with your Microsoft Tenant ID
const getAccessTokenLink = 'https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token'

// link used to create Teams online meeting
//replace with your Microsoft ID
const meetingCreationLink = 'https://graph.microsoft.com/v1.0/users/YOUR_MICROSOFT_ID/onlineMeetings'



// parameters used to get the access token
const requestParams = qs.stringify({
    grant_type: 'client_credentials',
    client_id: 'YOUR_CLIENT_ID',
    client_secret: 'YOUR_CLIENT_SECRET',
    scope: 'https://graph.microsoft.com/.default'
})

// parameters udes to create the meeting
// date format must be ISO 8601
// e.g. : 2023-08-23T20:30:00.0000000+02:00
const meetingCreationParams = JSON.stringify({
    "startDateTime": "START_DATE",
    "endDateTime": "END_DATE",
    "subject": "MS Teams meeting",
    "lobbyBypassSettings": {
        "isDialInBypassEnabled": true,
        "scope": "everyone"
    },
    "participants": {
        "organizer": {
            "identity": {
                "user": {
                    "id": "YOUR_MICROSOFT_ID"
                }
            }
        }
    }
})



bot.onText(/\/call/, async(msg) => {
   var accessToken = await getAccessToken()
   var teamsLink = await getTeamsLink(accessToken)

   // maybe use a URL shortener before sending the link ! ! !
   await bot.sendMessage(msg.chat.id, "Microsoft Teams link ⬇️\n\n" + teamsLink)
});

async function getAccessToken() {
    try {

        const response = await axios.post(getAccessTokenLink, requestParams)
        var parsedBody = response.data
        var accessToken = parsedBody.access_token

        return accessToken

    } catch (error) {

        console.log("Error while genereting acces token : " + error)

        // full server error response
        if (error.response) {
            console.error("Status code :", error.response.status)
            console.error("Date response :", error.response.data)
            console.error("Header response :", error.response.headers)
        }

        return
    }
}

async function getTeamsLink(accessToken){
    try {

        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${accessToken}`
        }

        const response = await axios.post(meetingCreationLink, meetingCreationParams, { headers })
        const teamsLink = response.data.joinWebUrl
        
        return teamsLink

    } catch (error) {

        console.error("Error while generating Teams link :", error.message)

        // full server error response
        if (error.response) {
            console.error("Status code :", error.response.status)
            console.error("Date response :", error.response.data)
            console.error("Header response :", error.response.headers)
        }

        return
    }
}
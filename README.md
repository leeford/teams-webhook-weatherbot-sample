# teams-webhook-weatherbot-sample
A simple Teams bot to get the weather using an outgoing webhook. The solution comprises of:

- Azure Functions HTTP Trigger (Bot)
- Outgoing Webhook in a Team

A user @ mentions the webhook and a location (for weather) and this triggers the Azure function, which calls the [OpenWeatherMap API](https://openweathermap.org/api) for the weather information. An Adaptive Card is returned with the information to the user.

![demo](https://user-images.githubusercontent.com/472320/149221643-96f50590-dd93-4616-83f5-98ff9d219a4b.gif)
## Prerequisites

- You will need to register for an API key from [OpenWeatherMap](https://openweathermap.org)
- You will need to create an Outgoing Webhook in a Team and make a copy of the security token

## Usage
Deploy functions in to Azure or run locally. 

*Run locally*
```bash
npm install
npm run start
```

You will need to populate the following application settings.

- **TeamsSecurityToken**: The **Security token** from the Teams Outgoing Webhook. This is used for the bot to validate requests are from a trusted source (your Team)
- **WeatherAPIKey**: The **API key** to grant access to the OpenWeatherMap API

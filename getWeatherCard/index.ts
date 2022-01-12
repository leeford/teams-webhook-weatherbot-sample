import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { Activity, CardFactory, MessageFactory } from "botbuilder";
import * as ACData from "adaptivecards-templating";
import axios from "axios";
import * as crypto from "crypto";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    // Check we have authorization header (should be HMAC) and body
    if (req.headers.authorization && req.body) {
        // Generate buffer from stored Teams Security Token
        const bufferToken = Buffer.from(process.env.TeamsSecurityToken, "base64");
        // Generate HMAC of message received with stored Teams Security Token
        const bufferMessage = Buffer.from(JSON.stringify(req.body), "utf-8");
        const generatedHMAC = "HMAC " + crypto.createHmac("sha256", bufferToken).update(bufferMessage).digest("base64");
        // Check that our generated HMAC matches with authorization header
        if (generatedHMAC === req.headers.authorization) {
            let message: Partial<Activity>;
            // Extract location name from text
            const locationName: string = req.body.text.replace(/<at>.*<\/at>(&nbsp;)?/, '').trim();
            // Get current weather for location
            await axios.get(`https://api.openweathermap.org/data/2.5/weather?q=${locationName}&appid=${process.env.WeatherApiKey}&units=metric`)
                .then(async currentWeatherResponse => {
                    const weatherData = currentWeatherResponse.data as any;

                    await axios.get(`https://api.openweathermap.org/data/2.5/onecall?lon=${weatherData.coord.lon}&lat=${weatherData.coord.lat}&appid=${process.env.WeatherApiKey}&exclude=minutely,hourly&units=metric`)
                        .then(onecallResponse => {

                            const oneCallData = onecallResponse.data as any;

                            // Create daily forecast data
                            const dailyForecastData: any[] = [];
                            oneCallData.daily.forEach((day: { dt: number, weather: { icon: string, main: string; }[]; temp: { max: number; min: number; }; }) => {
                                dailyForecastData.push({
                                    date: new Date((day.dt + oneCallData.timezone_offset) * 1000).toDateString(),
                                    imageUrl: `http://openweathermap.org/img/wn/${day.weather[0].icon}.png`,
                                    description: day.weather[0].main,
                                    high: day.temp.max.toFixed(),
                                    low: day.temp.min.toFixed()
                                })
                            });

                            // Create data for card
                            const cardData = {
                                location: `${weatherData.name}, ${weatherData.sys.country}`,
                                currentTemp: oneCallData.current.temp.toFixed(),
                                currentDescription: oneCallData.current.weather[0].main,
                                sunrise: new Date((oneCallData.current.sunrise + oneCallData.timezone_offset) * 1000).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                                sunset: new Date((oneCallData.current.sunset + oneCallData.timezone_offset) * 1000).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                                currentHumidity: oneCallData.current.humidity,
                                currentWindSpeed: oneCallData.current.wind_speed.toFixed(),
                                currentImageUrl: `http://openweathermap.org/img/wn/${oneCallData.current.weather[0].icon}@2x.png`,
                                dailyForecastData
                            }

                            // Card definition
                            const cardDefinition = {
                                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                                "type": "AdaptiveCard",
                                "version": "1.4",
                                "body": [
                                    {
                                        "type": "TextBlock",
                                        "text": "${location}",
                                        "size": "ExtraLarge",
                                        "isSubtle": true,
                                        "wrap": true
                                    },
                                    {
                                        "type": "ColumnSet",
                                        "columns": [
                                            {
                                                "type": "Column",
                                                "width": "102px",
                                                "items": [
                                                    {
                                                        "type": "Image",
                                                        "url": "${currentImageUrl}",
                                                        "width": "100px",
                                                        "height": "100px"
                                                    }
                                                ]
                                            },
                                            {
                                                "type": "Column",
                                                "width": 50,
                                                "items": [
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "${currentTemp}°C",
                                                        "size": "ExtraLarge",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    },
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "${currentDescription}",
                                                        "size": "Large",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    }
                                                ]
                                            },
                                            {
                                                "type": "Column",
                                                "width": 50,
                                                "items": [
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "Sunrise: ${sunrise}",
                                                        "horizontalAlignment": "Left",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    },
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "Sunset: ${sunset}",
                                                        "horizontalAlignment": "Left",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    },
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "Humidity: ${currentHumidity}%",
                                                        "horizontalAlignment": "Left",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    },
                                                    {
                                                        "type": "TextBlock",
                                                        "text": "Wind Speed: ${currentWindSpeed}mph",
                                                        "horizontalAlignment": "Left",
                                                        "spacing": "None",
                                                        "wrap": true
                                                    }
                                                ]
                                            }
                                        ]
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "Forecast",
                                        "wrap": true,
                                        "spacing": "Medium",
                                        "size": "Large"
                                    },
                                    {
                                        "type": "Container",
                                        "$data": "${dailyForecastData}",
                                        "items": [
                                            {
                                                "type": "ColumnSet",
                                                "columns": [
                                                    {
                                                        "type": "Column",
                                                        "width": "52px",
                                                        "items": [
                                                            {
                                                                "type": "Image",
                                                                "url": "${imageUrl}",
                                                                "width": "50px",
                                                                "height": "50px",
                                                                "horizontalAlignment": "Center"
                                                            }
                                                        ],
                                                        "verticalContentAlignment": "Center"
                                                    },
                                                    {
                                                        "type": "Column",
                                                        "width": 75,
                                                        "items": [
                                                            {
                                                                "type": "ColumnSet",
                                                                "columns": [
                                                                    {
                                                                        "type": "Column",
                                                                        "width": 50,
                                                                        "items": [
                                                                            {
                                                                                "type": "TextBlock",
                                                                                "text": "${date}",
                                                                                "wrap": true,
                                                                                "size": "Large",
                                                                                "spacing": "Small"
                                                                            },
                                                                            {
                                                                                "type": "TextBlock",
                                                                                "text": "${description}",
                                                                                "wrap": true,
                                                                                "spacing": "Small"
                                                                            }
                                                                        ]
                                                                    },
                                                                    {
                                                                        "type": "Column",
                                                                        "width": 50,
                                                                        "items": [
                                                                            {
                                                                                "type": "TextBlock",
                                                                                "text": "${high}°C",
                                                                                "wrap": true,
                                                                                "size": "Large",
                                                                                "spacing": "Small"
                                                                            },
                                                                            {
                                                                                "type": "TextBlock",
                                                                                "text": "${low}°C",
                                                                                "wrap": true,
                                                                                "spacing": "Small"
                                                                            }
                                                                        ]
                                                                    }
                                                                ],
                                                                "spacing": "None"
                                                            }
                                                        ],
                                                        "verticalContentAlignment": "Center"
                                                    }
                                                ]
                                            }
                                        ],
                                        "separator": true
                                    }
                                ]
                            }

                            const template = new ACData.Template(cardDefinition);
                            const cardPayload = template.expand({ $root: cardData });
                            const card = CardFactory.adaptiveCard(cardPayload);
                            message = MessageFactory.attachment(card);
                        });
                })
                .catch(error => {
                    if (error.response.status === 404) {
                        message = MessageFactory.text(`Sorry, I couldn't find weather for ${locationName}.`);
                    } else {
                        throw new Error("Sorry, an error occurred.");
                    }
                });

            context.res = {
                body: message
            };
        }
    }
};

export default httpTrigger;
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.BotBuilderSamples.Models;
using Microsoft.Bot.Connector.Authentication;
using System;
using TabWithAdpativeCardFlow;
using Microsoft.BotBuilderSamples.Helpers;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsTaskModuleBot : TeamsActivityHandler
    {
        private readonly string _baseUrl;
        private readonly string _connectionName;

        public TeamsTaskModuleBot(IConfiguration config)
        {
            _connectionName = config["ConnectionName"] ?? throw new NullReferenceException("ConnectionName");
            _baseUrl = config["BaseUrl"].EndsWith("/") ? config["BaseUrl"] : config["BaseUrl"] + "/";
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var reply = MessageFactory.Attachment(new[] { GetTaskModuleHeroCardOptions(), GetTaskModuleAdaptiveCardOptions() });
            await turnContext.SendActivityAsync(reply, cancellationToken);
        }

        /// <summary>
        /// Invoked when an fetch activity is recieved for tab.
        /// </summary>
        /// <param name="turnContext"></param>
        /// <param name="tabRequest"></param>
        /// <param name="cancellationToken"></param>
        /// <returns>Tab response.</returns>
        protected override async Task<TabResponse> OnTeamsTabFetchAsync(ITurnContext<IInvokeActivity> turnContext, TabRequest tabRequest, CancellationToken cancellationToken)
        {
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            if (tabRequest.TabEntityContext.TabEntityId == "homeTab")
            {

                // Check the state value
                var state = JsonConvert.DeserializeObject<AdaptiveCardAction>(turnContext.Activity.Value.ToString());
                var tokenResponse = await GetTokenResponse(turnContext, state.State, cancellationToken);

                if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
                {
                    // There is no token, so the user has not signed in yet.
                    var resource = await userTokenClient.GetSignInResourceAsync(_connectionName, turnContext.Activity as Activity, null, cancellationToken);

                    // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                    var signInLink = resource.SignInLink;

                    return new TabResponse
                    {
                        Tab = new TabResponsePayload
                        {
                            Type = "auth",
                            SuggestedActions = new TabSuggestedActions
                            {
                                Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = signInLink,
                                        Title = "Sign in to this app",
                                    },
                                },
                            },
                        },
                    };
                }

                var client = new SimpleGraphClient(tokenResponse.Token);
                var profile = await client.GetUserProfile();
                var userPhoto = await client.GetPublicURLForProfilePhoto(_baseUrl);

                return new TabResponse
                {
                    Tab = new TabResponsePayload
                    {
                        Type = "continue",
                        Value = new TabResponseCards
                        {
                            Cards = new List<TabResponseCard>
                            {
                                new TabResponseCard
                                {
                                    Card = CardHelper.GetSampleAdaptiveCard1(userPhoto, profile.DisplayName)
                                },
                                new TabResponseCard
                                {
                                    Card = CardHelper.GetSampleAdaptiveCard2()
                                },
                            },
                        },
                    },
                };
            }
            else
            {
                return new TabResponse
                {
                    Tab = new TabResponsePayload
                    {
                        Type = "continue",
                        Value = new TabResponseCards
                        {
                            Cards = new List<TabResponseCard>
                            {
                                new TabResponseCard
                                {
                                    Card = CardHelper.GetSampleAdaptiveCard3()
                                },
                            },
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Invoked when an submit activity is recieved for tab.
        /// </summary>
        /// <param name="turnContext"></param>
        /// <param name="tabSubmit"></param>
        /// <param name="cancellationToken"></param>
        /// <returns>Tab response.</returns>
        protected async override Task<TabResponse> OnTeamsTabSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TabSubmit tabSubmit, CancellationToken cancellationToken)
        {
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            await userTokenClient.SignOutUserAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, cancellationToken);

            return new TabResponse
            {
                Tab = new TabResponsePayload
                {
                    Type = "continue",
                    Value = new TabResponseCards
                    {
                        Cards = new List<TabResponseCard>
                            {
                                new TabResponseCard
                                {
                                    Card = CardHelper.GetSignOutCard()
                                },
                            },
                    },
                },
            };
        }

        protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var taskInfo = new TaskModuleResponse();
            var asJobject = JObject.FromObject(taskModuleRequest.Data);
            var buttonType = (string)asJobject.ToObject<CardTaskFetchValue<string>>()?.Id;

            if (buttonType == "btntypevalue1")
            {
               // var videoId = asJobject.GetValue("youTubeVideoId")?.ToString();
                taskInfo.Task = new TaskModuleContinueResponse
                {
                    Type = "continue",
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = _baseUrl + TaskModuleIds.CustomForm,
                        Height = 1000,
                        Width = 700,
                        Title = "Lookup a coworker",
                    },
                };
            }
            else
            {
                taskInfo.Task = new TaskModuleContinueResponse
                {
                    Type = "continue",
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = CardHelper.GetAdaptiveCardForTaskModule(),
                        Height = 200,
                        Width = 350,
                        Title = "Sample Adaptive Card",
                    },
                };
            }

            return Task.FromResult(taskInfo);
            /*
            var asJobject = JObject.FromObject(taskModuleRequest.Data);
            var value = asJobject.ToObject<CardTaskFetchValue<string>>()?.Data;

            var taskInfo = new TaskModuleTaskInfo();
            switch (value)
            {
                case TaskModuleIds.YouTube:
                    taskInfo.Url = taskInfo.FallbackUrl = _baseUrl + TaskModuleIds.YouTube;
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.YouTube);
                    break;
                case TaskModuleIds.CustomForm:
                    taskInfo.Url = taskInfo.FallbackUrl = _baseUrl + TaskModuleIds.CustomForm;
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.CustomForm);
                    break;
                case TaskModuleIds.AdaptiveCard:
                    taskInfo.Card = CreateAdaptiveCardAttachment();
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.AdaptiveCard);
                    break;
                default:
                    break;
            }

            return Task.FromResult(taskInfo.ToTaskModuleResponse());*/
        }

        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var reply = MessageFactory.Text("OnTeamsTaskModuleSubmitAsync Value: " + JsonConvert.SerializeObject(taskModuleRequest));
            await turnContext.SendActivityAsync(reply, cancellationToken);

            return TaskModuleResponseFactory.CreateResponse("Thanks!");
        }

        private static void SetTaskInfo(TaskModuleTaskInfo taskInfo, UISettings uIConstants)
        {
            taskInfo.Height = uIConstants.Height;
            taskInfo.Width = uIConstants.Width;
            taskInfo.Title = uIConstants.Title.ToString();
        }

        private static Attachment GetTaskModuleHeroCardOptions()
        {
            // Create a Hero Card with TaskModuleActions for each Task Module
            return new HeroCard()
            {
                Title = "Task Module Invocation from Hero Card - Testing YMAL",
                Buttons = new[] { TaskModuleUIConstants.AdaptiveCard, TaskModuleUIConstants.CustomForm, TaskModuleUIConstants.YouTube }
                            .Select(cardType => new TaskModuleAction(cardType.ButtonTitle, new CardTaskFetchValue<string>() { Data = cardType.Id }))
                            .ToList<CardAction>(),
            }.ToAttachment();
        }

        private static Attachment GetTaskModuleAdaptiveCardOptions()
        {
            // Create an Adaptive Card with an AdaptiveSubmitAction for each Task Module
            var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>()
                    {
                        new AdaptiveTextBlock(){ Text="Task Module Invocation from Adaptive Card", Weight=AdaptiveTextWeight.Bolder, Size=AdaptiveTextSize.Large}
                    },
                Actions = new[] { TaskModuleUIConstants.AdaptiveCard, TaskModuleUIConstants.CustomForm, TaskModuleUIConstants.YouTube }
                            .Select(cardType => new AdaptiveSubmitAction() { Title = cardType.ButtonTitle, Data = new AdaptiveCardTaskFetchValue<string>() { Data = cardType.Id } })
                            .ToList<AdaptiveAction>(),
            };

            return new Attachment() { ContentType = AdaptiveCard.ContentType, Content = card };
        }

        private static Attachment CreateAdaptiveCardAttachment()
        {
            // combine path for cross platform support
            string[] paths = { ".", "Resources", "adaptiveCard.json" };
            var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Get token response on basis of state.
        /// </summary>
        private async Task<TokenResponse> GetTokenResponse(ITurnContext<IInvokeActivity> turnContext, string state, CancellationToken cancellationToken)
        {
            var magicCode = string.Empty;

            if (!string.IsNullOrEmpty(state))
            {
                if (int.TryParse(state, out var parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            var tokenResponse = await userTokenClient.GetUserTokenAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, magicCode, cancellationToken).ConfigureAwait(false);
            return tokenResponse;
        }
    }
}

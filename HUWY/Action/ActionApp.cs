using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using AdaptiveCards.Templating;

namespace HUWY.Action;

public class ActionApp : TeamsActivityHandler
{
    //private readonly string _adaptiveCardFilePath = Path.Combine(".", "Resources", "helloWorldCard.json");

    //gavdcodebegin 001
    protected override async Task<MessagingExtensionActionResponse>
                                OnTeamsMessagingExtensionSubmitActionAsync(
                                        ITurnContext<IInvokeActivity> turnContext,
                                        MessagingExtensionAction action,
                                        CancellationToken cancellationToken)
    {
        CardResponse actionData = ((JObject)action.Data).ToObject<CardResponse>();

        int intNumberOne = int.Parse(actionData.NumberOne);
        int intNumberTwo = int.Parse(actionData.NumberTwo);
        int intSum = intNumberOne + intNumberTwo;

        string sumCardFilePath = Path.Combine(".", "Resources", "sumCard.json");

        string templateJson = await File.ReadAllTextAsync(
                                            sumCardFilePath, cancellationToken);
        AdaptiveCardTemplate template = new(templateJson);
        string adaptiveCardJson = template.Expand(new
        {
            title = "The sum of " + actionData.NumberOne +
                                            " and " + actionData.NumberTwo + " is:",
            sum = intSum.ToString()
        });

        AdaptiveCard adaptiveCard = AdaptiveCard.FromJson(adaptiveCardJson).Card;
        MessagingExtensionAttachment attachments = new()
        {
            ContentType = AdaptiveCard.ContentType,
            Content = adaptiveCard
        };

        return new MessagingExtensionActionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = new[] { attachments }
            }
        };
    }
    //gavdcodeend 001
}


//gavdcodebegin 002
internal class CardResponse
{
    public string Title { get; set; }
    public string SubTitle { get; set; }
    public string Text { get; set; }
    public string NumberOne { get; set; }
    public string NumberTwo { get; set; }
}
//gavdcodeend 002

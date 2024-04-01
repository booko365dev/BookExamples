using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using AdaptiveCards.Templating;
using System.Net.Http.Headers;

namespace DSPG.Search;

public class SearchApp : TeamsActivityHandler
{
    private readonly string _adaptiveCardFilePath = Path.Combine(
                                            ".", "Resources", "helloWorldCard.json");

    //gavdcodebegin 001
    protected override async Task<MessagingExtensionResponse>
                                OnTeamsMessagingExtensionQueryAsync(
                                            ITurnContext<IInvokeActivity> turnContext,
                                            MessagingExtensionQuery query,
                                            CancellationToken cancellationToken)
    {
        string templateJson = await File.ReadAllTextAsync(
                                            _adaptiveCardFilePath, cancellationToken);

        string userQuery = query?.Parameters?[0]?.Value as string ?? string.Empty;

        Wikipedia wikiResults = await GetWikipediaResult(userQuery);
        string wikiUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/" +
                    "Orange_Icon_Pale_Wiki.png/120px-Orange_Icon_Pale_Wiki.png";

        AdaptiveCardTemplate myCardTemplate = new(templateJson);
        List<MessagingExtensionAttachment> messAllAttachments = [];

        (AdaptiveCard, ThumbnailCard) myCard = CreateCard(wikiResults, wikiUrl, myCardTemplate);

        MessagingExtensionAttachment messOneAttachment = new()
        {
            ContentType = AdaptiveCard.ContentType,
            Content = myCard.Item1,
            Preview = myCard.Item2.ToAttachment()
        };
        messAllAttachments.Add(messOneAttachment);

        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = messAllAttachments
            }
        };
    }
    //gavdcodeend 001

    //gavdcodebegin 002
    private static async Task<Wikipedia> GetWikipediaResult(string WordToQuery)
    {
        HttpClient client = new HttpClient();

        string wikiUrl = "https://en.wikipedia.org/w/api.php?" +
            "action=query&list=search&srsearch=" + WordToQuery +
            "&utf8=&format=json";

        client.BaseAddress = new Uri(wikiUrl);
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(
        new MediaTypeWithQualityHeaderValue("application/json"));

        Wikipedia myResult = null;
        HttpResponseMessage response = await client.GetAsync(client.BaseAddress);
        if (response.IsSuccessStatusCode)
        {
            string myResultJson = await response.Content.ReadAsStringAsync();
            myResult = System.Text.Json.JsonSerializer.Deserialize<Wikipedia>(
                                                                        myResultJson);
        }

        return myResult;
    }
    //gavdcodeend 002

    //gavdcodebegin 004
    private static (AdaptiveCard, ThumbnailCard) CreateCard(
                                                    Wikipedia wikiResults,
                                                    string wikiUrl,
                                                    AdaptiveCardTemplate myCardTemplate)
    {
        ThumbnailCard wikiPreviewCard = new ThumbnailCard
        {
            Title = wikiResults.query.search[0].title
        };

        string adaptiveCardJson = myCardTemplate.Expand(new
        {
            name = wikiResults.query.search[0].title,
            description = wikiResults.query.search[0].snippet
        });
        AdaptiveCard wikiAdaptiveCard = AdaptiveCard.FromJson(adaptiveCardJson).Card;
        if (!string.IsNullOrEmpty(wikiUrl))
        {
            wikiPreviewCard.Images = new List<CardImage>() { new(wikiUrl, "Icon") };
            wikiAdaptiveCard.Body.Insert(0, new AdaptiveImage()
            {
                Url = new Uri(wikiUrl),
                Style = AdaptiveImageStyle.Person,
                Size = AdaptiveImageSize.Small,
            });
        }

        return (wikiAdaptiveCard, wikiPreviewCard);
    }
    //gavdcodeend 004
}

//gavdcodebegin 003
public class Wikipedia
{
    public string batchcomplete { get; set; }
    public Continue _continue { get; set; }
    public Query query { get; set; }
}

public class Continue
{
    public int sroffset { get; set; }
    public string _continue { get; set; }
}

public class Query
{
    public Searchinfo searchinfo { get; set; }
    public Search[] search { get; set; }
}

public class Searchinfo
{
    public int totalhits { get; set; }
}

public class Search
{
    public int ns { get; set; }
    public string title { get; set; }
    public int pageid { get; set; }
    public int size { get; set; }
    public int wordcount { get; set; }
    public string snippet { get; set; }
    public DateTime timestamp { get; set; }
}
//gavdcodeend 003

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using System.Net.Http.Headers;

namespace QEXM.Bot;

public class EchoBot : TeamsActivityHandler
{
    //gavdcodebegin 001
    protected override async Task OnMessageActivityAsync(
                                ITurnContext<IMessageActivity> turnContext,
                                CancellationToken cancellationToken)
    {
        string messageText = turnContext.Activity.RemoveRecipientMention()?.Trim();

        string aboutText = GetWordAfterAbout(messageText);
        string wikipediaText = await GetWikipediaSnippet(aboutText);
        string replyText = $"My Wikipedia bot says: {wikipediaText}";

        await turnContext.SendActivityAsync(
                                MessageFactory.Text(replyText),
                                cancellationToken);
    }
    //gavdcodeend 001

    protected override async Task OnMembersAddedAsync(
                                IList<ChannelAccount> membersAdded,
                                ITurnContext<IConversationUpdateActivity> turnContext,
                                CancellationToken cancellationToken)
    {
        var welcomeText = "Hi there! I'm a Teams bot to answer all your questions.";
        foreach (var member in membersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(
                                MessageFactory.Text(welcomeText),
                                cancellationToken);
            }
        }
    }

    //gavdcodebegin 002
    static string GetWordAfterAbout(string QueryText)
    {
        string[] wordsInText = QueryText.Split(' ');
        string afterAbout = string.Empty;
        for (int myCounter = 0; myCounter < wordsInText.Length; myCounter++)
        {
            if (wordsInText[myCounter].Trim().Equals(
                                    "about", StringComparison.CurrentCultureIgnoreCase))
            {
                afterAbout = wordsInText[myCounter + 1];
                break;
            }
        }

        return afterAbout;
    }
    //gavdcodeend 002

    //gavdcodebegin 003
    static async Task<string> GetWikipediaSnippet(string WordToQuery)
    {
        HttpClient client = new();

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

        string strReturn = myResult.query.search[0].snippet;
        return strReturn;
    }
    //gavdcodeend 003
}

//gavdcodebegin 004
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
//gavdcodeend 004


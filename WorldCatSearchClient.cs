using System.Text;
using Flurl.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

class WorldCatSearchClient
{
    const string OAuthUrl = "https://oauth.oclc.org/token";
    const string BaseUrl = "https://americas.discovery.api.oclc.org/worldcat/search/v2";
    private FlurlClient Client { get; }
    private string ClientId { get; }
    private string ClientSecret { get; }
    private string AccessToken { get; set; }
    private DateTimeOffset AccessTokenExpiresAt { get; set; }

    public WorldCatSearchClient(string clientId, string clientSecret)
    {
        ClientId = clientId;
        ClientSecret = clientSecret;
        Client = new FlurlClient(BaseUrl).BeforeCall(call => Authorize(call.Request));
    }

    private async Task Authorize(IFlurlRequest request)
    {
        try
        {
            if (AccessToken is null || AccessTokenExpiresAt.AddSeconds(-5) < DateTimeOffset.Now)
            {
                var authResponse = await OAuthUrl
                    .WithBasicAuth(ClientId, ClientSecret)
                    .PostUrlEncodedAsync(new 
                    {
                        grant_type = "client_credentials",
                        scope = "wcapi:view_bib wcapi:view_brief_bib wcapi:view_retained_holdings wcapi:view_summary_holdings wcapi:view_my_holdings wcapi:view_institution_holdings wcapi:view_holdings",
                    })
                    .ReceiveString();
                var authJson = JsonConvert.DeserializeObject<JObject>(authResponse);
                AccessToken = authJson["access_token"].ToString();
                AccessTokenExpiresAt = DateTimeOffset.Parse(authJson["expires_at"].ToString());
            }
        }
        catch (FlurlHttpException e)
        {
            var response = await e.GetResponseStringAsync();
            throw;
        }
        request.WithHeader("Authorization", "Bearer " + AccessToken);
    }

    public async Task<JObject> GetHoldings(string oclcControlNumber)
    {
        try
        {
            var response = await Client.Request("bibs-holdings").SetQueryParam("oclcNumber", oclcControlNumber).GetStringAsync();
            var responseJson = JsonConvert.DeserializeObject<JObject>(response);
            return responseJson;
        }
        catch (FlurlHttpException e)
        {
            var response = await e.GetResponseStringAsync();
            throw;
        }
    }
}

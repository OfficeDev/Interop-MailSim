using MailSim.Common;
using System.Web;
using System.Collections.Generic;


namespace MailSim.ProvidersREST
{
    /// <summary>
    /// Provides clients for the different service endpoints.
    /// </summary>
    internal static class AuthenticationHelperHTTP
    {
        private static readonly string ClientID = Resources.ClientID;
        private const string AadServiceResourceId = "https://graph.windows.net/";

        // Properties used for communicating with your Windows Azure AD tenant.
        private const string CommonAuthority = "https://login.microsoftonline.com/Common";
 
        private static string UserName { get; set; }
        private static string Password { get; set; }

        private class AccessTokenResponse
        {
            public string access_token { get; set; }
            public int expires_in { get; set; }
            public int expires_on { get; set; }
            public string id_token { get; set; }
            public string refresh_token { get; set; }
            public string resource { get; set; }
            public string scope { get; set; }
            public string token_type { get; set; }
        }

        private static IDictionary<string, AccessTokenResponse> _tokenResponses = new Dictionary<string, AccessTokenResponse>();

        internal static void Initialize(string userName, string password)
        {
            if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password))
            {
                throw new System.Exception("No user name or password in config file!");
            }

            UserName = userName;
            Password = password;
        }

        private static string TokenUri
        {
            get
            {
                return CommonAuthority + "/oauth2/" + "token";
            }
        }

        private static string GetTokenHelperHttp(string resourceId, bool isRefresh)
        {
            if (isRefresh)
            {
                var authResponse = _tokenResponses[resourceId];

                string body = string.Format("grant_type=refresh_token&refresh_token={0}&client_id={1}&resource={2}",
                                                HttpUtility.UrlEncode(authResponse.refresh_token),
                                                HttpUtility.UrlEncode(ClientID),
                                                HttpUtility.UrlEncode(resourceId)
                                                );

                Log.Out(Log.Severity.Info, "", "Sending request for new token:" + body);

                var newAuthResponse = DoTokenHttp(body);

                _tokenResponses[resourceId] = newAuthResponse;

                Log.Out(Log.Severity.Info, "", "Got new access token:" + newAuthResponse.access_token);
            }

            if (_tokenResponses.ContainsKey(resourceId) == false)
            {
                _tokenResponses[resourceId] = QueryTokenResponse(resourceId);
            }

            return _tokenResponses[resourceId].access_token;
        }

        private static AccessTokenResponse QueryTokenResponse(string resourceId)
        {
            string body = string.Format("resource={0}&client_id={1}&grant_type=password&username={2}&password={3}&scope=openid",
                                            HttpUtility.UrlEncode(resourceId),
                                            HttpUtility.UrlEncode(ClientID),
                                            HttpUtility.UrlEncode(UserName),
                                            HttpUtility.UrlEncode(Password));

            return DoTokenHttp(body);
        }

        private static AccessTokenResponse DoTokenHttp(string body)
        {
            return HttpUtil.DoHttp<string, AccessTokenResponse>("POST", TokenUri, body, (dummy) => null).GetResult();
        }

        internal static string GetToken(string resourceId, bool isRefresh)
        {
            return GetTokenHelper(resourceId, isRefresh);
        }

        private static string GetTokenHelper(string resourceId, bool isRefresh)
        {
            return GetTokenHelperHttp(resourceId, isRefresh);
        }
    }
}

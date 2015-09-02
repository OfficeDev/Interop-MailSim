using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Net.Http.Formatting;
using System.Dynamic;
using System.Linq;

namespace MailSim.ProvidersREST
{
    static class HttpUtil
    {
        private const string baseOutlookUri = @"https://outlook.office.com/api/v1.0/Me/";

        internal static async Task<T> GetItemAsync<T>(string uri)
        {
            return await DoHttp<T,T>(HttpMethod.Get, uri, default(T));
        }

        internal static async Task<T> GetItemsAsync<T>(string uri)
        {
            var coll = await GetCollectionAsync<T>(uri);

            return coll.value;
        }

        internal static IEnumerable<T> EnumerateCollection<T>(string uri, int count)
        {
#if false
            IEnumerable<T> items = Enumerable.Empty<T>();

            while (count > 0 && uri != null)
            {
                var msgsColl = GetCollectionAsync<IEnumerable<T>>(uri).Result;

                items = items.Union(msgsColl.value.Take(count));
                count -= msgsColl.value.Count();

                uri = msgsColl.NextLink;
            }

            return items;
#else
            while (uri != null)
            {
                var msgsColl = GetCollectionAsync<IEnumerable<T>>(uri).Result;

                foreach (var m in msgsColl.value)
                {
                    if (--count <= 0)
                    {
                        yield break;
                    }
                    yield return m;
                }

                uri = msgsColl.NextLink;
            }
#endif
        }

        internal static async Task<ODataCollection<T>> GetCollectionAsync<T>(string uri)
        {
            return await DoHttp<ODataCollection<T>, ODataCollection<T>>(HttpMethod.Get, uri, default(ODataCollection<T>));
        }

        internal static async Task<T> PostItemAsync<T>(string uri, T item=default(T))
        {
            return await DoHttp<T, T>(HttpMethod.Post, uri, item);
        }

        internal static async Task<T> PostDynamicAsync<T>(string uri, dynamic body)
        {
            return await DoHttp<ExpandoObject, T>(HttpMethod.Post, uri, body);
        }

        internal static async Task DeleteAsync(string uri)
        {
            using (HttpClient client = GetHttpClient())
            {
                var response = await client.DeleteAsync(BuildUri(uri));

                response.EnsureSuccessStatusCode();
            }
        }

        internal static async Task<T> PatchItemAsync<T>(string uri, T item)
        {
            return await DoHttp<T,T>("PATCH", uri, item);
        }

        private static async Task<TResult> DoHttp<TBody, TResult>(HttpMethod method, string uri, TBody body)
        {
            HttpResponseMessage response;

            using (HttpClient client = GetHttpClient())
            {
                var request = new HttpRequestMessage(method, BuildUri(uri));

                if (EqualityComparer<TBody>.Default.Equals(body, default(TBody)) == false)
                {
                    request.Content = new ObjectContent<TBody>(body, new JsonMediaTypeFormatter());
                }

                response = await client.SendAsync(request);
            }

            string jsonResponse = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<TResult>(jsonResponse);
            }
            else
            {
                var errorDetail = JsonConvert.DeserializeObject<ODataError>(jsonResponse);
                throw new System.Exception(errorDetail.error.message);
            }
        }

        private class ODataError
        {
            public class Error
            {
                public string code { get; set; }
                public string message { get; set; }
            }
            public Error error { get; set; }
        }

        internal static async Task<TResult> DoHttp<TBody, TResult>(string methodName, string uri, TBody body)
        {
            return await DoHttp<TBody,TResult>(new HttpMethod(methodName), uri, body);
        }

        private static string BuildUri(string subUri)
        {
            if (subUri.StartsWith("http"))
            {
                return subUri;
            }

            return baseOutlookUri + subUri;
        }

        private static HttpClient GetHttpClient()
        {
            HttpClient client = new HttpClient();

            string token = AuthenticationHelper.GetOutlookToken();

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            return client;
        }

        internal class ODataCollection<TCollection>
        {
            [Newtonsoft.Json.JsonProperty("@odata.nextLink")]
            public string NextLink { get; set; }

            public TCollection value { get; set; }
        }
    }
}

using Microsoft_Teams_Graph_RESTAPIs_Connect.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles
{
    public class ServiceHelper
    {
        public static string RedirectUri
        {
            get
            {
                return ConfigurationManager.AppSettings["ida:RedirectUri"];
            }
        }

        public static string AppId
        {
            get
            {
                return ConfigurationManager.AppSettings["ida:AppId"];
            }
        }

        public static string AppSecret
        {
            get
            {
                return ConfigurationManager.AppSettings["ida:AppSecret"];
            }
        }

        public static string Scopes
        {
            get
            {
                return ConfigurationManager.AppSettings["ida:GraphScopes"];
            }
        }


        //public static async Task<HttpResponseMessage> SendRequest(HttpMethod method, String endPoint, string accessToken, dynamic content = null)
        public static async Task<HttpResponseMessage> SendRequest(HttpMethod method, String endPoint, string accessToken, dynamic content = null)
        {
            HttpResponseMessage response = null;
            try
            {
                using (var client = new HttpClient())
                {
                    using (var request = new HttpRequestMessage(method, endPoint))
                    {
                        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        if (content != null)
                        {
                            string c;
                            if (content is string)
                                c = content;
                            else
                                c = JsonConvert.SerializeObject(content);
                            request.Content = new StringContent(c, Encoding.UTF8, "application/json");
                        }

                        //response = client.SendAsync(request).Result;
                        //response = await client.SendAsync(request);
                        response = await client.SendAsync(request).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            return response;
        }

        private static HttpClient outClient = new HttpClient();

        public static async Task<HttpResponseMessage> SendRequestTest(HttpMethod method, String endPoint, string accessToken, dynamic content = null)
        {
            try
            {
                HttpRequestMessage request = new HttpRequestMessage(method, endPoint);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                if (content != null)
                {
                    string c;
                    if (content is string)
                        c = content;
                    else
                        c = JsonConvert.SerializeObject(content);
                    request.Content = new StringContent(c, Encoding.UTF8, "application/json");
                }

                var response = await outClient.SendAsync(request).ConfigureAwait(false);
                return response;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }            
        }

        /// <summary>
        /// Helper function to prepare the ResultsItem list from request response.
        /// </summary>
        /// <param name="response">Request response</param>
        /// <param name="idPropertyName">Property name of the item Id</param>
        /// <param name="displayPropertyName">Property name of the item display name</param>
        /// <returns></returns>
        public static async Task<List<ResultsItem>> GetResultsItem(
            HttpResponseMessage response, string idPropertyName, string displayPropertyName, string resourcePropId)
        {
            List<ResultsItem> items = new List<ResultsItem>();

            JObject json = JObject.Parse(await response.Content.ReadAsStringAsync());
            foreach (JProperty content in json.Children<JProperty>())
            {
                if (content.Name.Equals("value"))
                {
                    var res = content.Value.AsJEnumerable().GetEnumerator();
                    res.MoveNext();

                    while (res.Current != null)
                    {
                        string display = "";
                        string id = "";

                        foreach (JProperty prop in res.Current.Children<JProperty>())
                        {
                            if (prop.Name.Equals(idPropertyName))
                            {
                                id = prop.Value.ToString();
                            }

                            if (prop.Name.Equals(displayPropertyName))
                            {
                                display = prop.Value.ToString();
                            }
                        }

                        items.Add(new ResultsItem
                        {
                            Display = display,
                            Id = id,
                            Properties = new Dictionary<string, object>
                                            {
                                                { resourcePropId, id }
                                            }
                        });

                        res.MoveNext();
                    }
                }
            }

            return items;
        }

    }
}
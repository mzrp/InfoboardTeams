using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft_Teams_Graph_RESTAPIs_Connect.Models;
using Newtonsoft.Json;
using Resources;
using System.Configuration;

using System.Net;
using System.IO;

namespace Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles
{
    public static class Statics
    {
        public static T Deserialize<T>(this string result)
        {
            return JsonConvert.DeserializeObject<T>(result);
        }
    }

    public partial class MsAuthToken
    {
        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("expires_in")]
        public long ExpiresIn { get; set; }

        [JsonProperty("ext_expires_in")]
        public long ExtExpiresIn { get; set; }

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }

    public class GraphService : HttpHelpers
    {
        private string GraphRootUri = ConfigurationManager.AppSettings["ida:GraphRootUri"];

        /// <summary>
        /// Create new channel.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <param name="teamId">Id of the team in which new channel needs to be created</param>
        /// <param name="channelName">New channel name</param>
        /// <param name="channelDescription">New channel description</param>
        /// <returns></returns>
        public async Task CreateChannel(string accessToken, string teamId, string channelName, string channelDescription)
        {
            await HttpPost($"/teams/{teamId}/channels",
                new Channel()
                {
                    description = channelDescription,
                    displayName = channelName
                });
        }

        public async Task<IEnumerable<Channel>> GetChannels(string accessToken, string teamId)
        {
            string endpoint = $"{GraphRootUri}/teams/{teamId}/channels";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
            return await ParseList<Channel>(response);
        }

        public async Task<IEnumerable<TeamsApp>> GetApps(string accessToken, string teamId)
        {
            // to do: switch to the V1 installedApps API
            return await HttpGetList<TeamsApp>($"/teams/{teamId}/apps", endpoint: graphBetaEndpoint);
        }

        public async Task<string> GetLoggedUserPresence(String accessToken)
        {
            // get logged user presence
            string sMyPresence = await GetMyPresence(accessToken);
            string sMyPresenceAvailability = sMyPresence.Split(',')[0];
            string sMyPresenceActivity = sMyPresence.Split(',')[1];
            System.Web.HttpContext.Current.Session["myPresence"] = sMyPresenceAvailability;
            return sMyPresenceAvailability;
        }

        /// <summary>
        /// Get the current users' ids.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<List<string>> GetUsersPresences(String accessToken, List<string> sAllTeamsUsers)
        {
            List<string> sAllTeamsUsersNew = sAllTeamsUsers;

            // get all teams users presences
            for (int i = 0; i < sAllTeamsUsersNew.Count; i++)
            {
                if (sAllTeamsUsersNew[i] != "")
                {
                    string[] sUserSipDataArray = sAllTeamsUsersNew[i].Split(',');
                    string sUserSip = sUserSipDataArray[0];
                    string sSystemId = sUserSipDataArray[1];
                    string sUserId = sUserSipDataArray[2];
                    string sId = sUserSipDataArray[3];
                    string sPresence = sUserSipDataArray[4];

                    if ((sUserSip == "<LOGGEDUSER>") && (sPresence == "<PRESENCE>"))
                    {
                        // get logged user presence
                        string sMyPresence = await GetMyPresence(accessToken);
                        string sMyPresenceAvailability = sMyPresence.Split(',')[0];
                        string sMyPresenceActivity = sMyPresence.Split(',')[1];
                        System.Web.HttpContext.Current.Session["myPresence"] = sMyPresenceAvailability;
                        sAllTeamsUsersNew[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + "," + sMyPresenceAvailability;
                    }

                    if ((sUserSip != "") && (sId == "<ID>") && (sPresence == "<PRESENCE>"))
                    {
                        // get id
                        string endpoint = "https://graph.microsoft.com/beta/users/" + sUserSip;
                        HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
                        if (response != null && response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            string userid = json.GetValue("id").ToString();

                            if (userid != null)
                            {
                                sId = userid;
                                sAllTeamsUsersNew[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + ",<PRESENCE>";
                            }
                        }
                    }

                    if ((sUserSip != "") && (sId != "<ID>") && (sPresence == "<PRESENCE>"))
                    {
                        string endpoint = "https://graph.microsoft.com/beta/users/" + sId + "/presence";
                        HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
                        if (response != null && response.IsSuccessStatusCode)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            string userpresence = json.GetValue("availability").ToString();
                            if (userpresence != null)
                            {
                                sPresence = userpresence;
                                sAllTeamsUsersNew[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + "," + sPresence;
                            }
                        }
                    }
                }
            }

            return sAllTeamsUsersNew;
        }

        /// <summary>
        /// Get the current users' ids.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<List<string>> GetUsersPresencesSingleFile(String accessToken)
        {
            List<string> sAllTeamsUsers = (List<string>)System.Web.HttpContext.Current.Session["sesAllTeamsUsers"];

            // get all teams users presences
            if (sAllTeamsUsers.Count > 0)
            {
                for (int i=0; i< sAllTeamsUsers.Count; i++)
                {
                    if (sAllTeamsUsers[i] != "")
                    {
                        string[] sUserSipDataArray = sAllTeamsUsers[i].Split(',');
                        string sUserSip = sUserSipDataArray[0];
                        string sSystemId = sUserSipDataArray[1];
                        string sUserId = sUserSipDataArray[2];
                        string sId = sUserSipDataArray[3];
                        string sPresence = sUserSipDataArray[4];

                        if ((sUserSip == "<LOGGEDUSER>") && (sPresence == "<PRESENCE>"))
                        {
                            // get logged user presence
                            string sMyPresence = await GetMyPresence(accessToken);
                            string sMyPresenceAvailability = sMyPresence.Split(',')[0];
                            string sMyPresenceActivity = sMyPresence.Split(',')[1];
                            System.Web.HttpContext.Current.Session["myPresence"] = sMyPresenceAvailability;
                            sAllTeamsUsers[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + "," + sMyPresenceAvailability;
                            return sAllTeamsUsers;
                        }

                        if ((sUserSip != "") && (sId == "<ID>") && (sPresence == "<PRESENCE>"))
                        {
                            // get id
                            string endpoint = "https://graph.microsoft.com/beta/users/" + sUserSip;
                            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
                            if (response != null && response.IsSuccessStatusCode)
                            {
                                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                                string userid = json.GetValue("id").ToString();

                                if (userid != null)
                                {
                                    sId = userid;
                                    sAllTeamsUsers[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + ",<PRESENCE>";
                                    return sAllTeamsUsers;
                                }
                            }
                        }

                        if ((sUserSip != "") && (sId != "<ID>") && (sPresence == "<PRESENCE>"))
                        {
                            string endpoint = "https://graph.microsoft.com/beta/users/" + sId + "/presence";
                            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
                            if (response != null && response.IsSuccessStatusCode)
                            {
                                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                                string userpresence = json.GetValue("availability").ToString();
                                if (userpresence != null)
                                {
                                    sPresence = userpresence;
                                    sAllTeamsUsers[i] = sUserSip + "," + sSystemId + "," + sUserId + "," + sId + "," + sPresence;
                                    return sAllTeamsUsers;
                                }
                            }
                        }
                    }
                }
            }

            return sAllTeamsUsers;
        }

        /// <summary>
        /// Get the current user's presence.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetUserId(String accessToken, string sSip)
        {
            string endpoint = "https://graph.microsoft.com/beta/users/" + sSip;
            String userid = "n/a";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
            if (response != null && response.IsSuccessStatusCode)
            {
                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                userid = json.GetValue("id").ToString();
            }
            return userid?.Trim();
        }

        /// <summary>
        /// Get Auth TOken
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetAuthToken()
        {
            string sResult = "n/a";

            try
            {
                var webRequestAUTH = WebRequest.Create("https://login.microsoftonline.com/74df0893-eb0e-4e6e-a68a-c5ddf3001c1f/oauth2/v2.0/token") as HttpWebRequest;
                if (webRequestAUTH != null)
                {
                    webRequestAUTH.Method = "POST";
                    webRequestAUTH.Host = "login.microsoftonline.com";
                    webRequestAUTH.ContentType = "application/x-www-form-urlencoded";

                    string sParams = "client_id=db772450-d87c-4f49-8296-403b8c4c4f19&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=IpA3trz3QUj16EAwRhutazOHpbW:r_-=&grant_type=client_credentials";
                    var data = Encoding.ASCII.GetBytes(sParams);
                    webRequestAUTH.ContentLength = data.Length;

                    using (var sW = webRequestAUTH.GetRequestStream())
                    {
                        sW.Write(data, 0, data.Length);
                    }

                    using (var rW = webRequestAUTH.GetResponse().GetResponseStream())
                    {
                        using (var srW = new StreamReader(rW))
                        {
                            var sExportAsJson = srW.ReadToEnd();
                            var sExport = JsonConvert.DeserializeObject<MsAuthToken>(sExportAsJson);

                            string sAuthToken = sExport.AccessToken;
                            string sTokenType = sExport.TokenType;
                            long lExpiresIn = sExport.ExpiresIn;

                            DateTime AuthTokenExpireIn = DateTime.Now;
                            AuthTokenExpireIn.AddSeconds(lExpiresIn);

                            System.Web.HttpContext.Current.Session["AuthToken"] = sAuthToken;
                            System.Web.HttpContext.Current.Session["AuthTokenType"] = sTokenType;
                            System.Web.HttpContext.Current.Session["AuthTokenExpireIn"] = AuthTokenExpireIn;

                            sResult = sAuthToken;
                        }
                    }

                    webRequestAUTH = null;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            return sResult;
        }

        /// <summary>
        /// Get report - users team details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsActivityCounts(String accessToken)
        {
            string sResult = "n/a";

            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 55)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get report - users team details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsDeviceUsage(String accessToken)
        {
            string sResult = "n/a";

            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 55)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get report - users team details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsDeviceCounts(String accessToken)
        {
            string sResult = "n/a";

            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserCounts(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 55)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get report - users team details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsDeviceDistribution(String accessToken)
        {
            string sResult = "n/a";

            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageDistributionUserCounts(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 55)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get report - users team counts.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsUserCounts(String accessToken)
        {
            string sResult = "n/a";

            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 55)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }


        /// <summary>
        /// Get report - users team details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetTeamsUserActivity(String accessToken)
        {
            string sResult = "n/a";
            
            try
            {
                if (accessToken != "")
                {
                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')") as HttpWebRequest;
                    if (webRequestTR != null)
                    {
                        webRequestTR.Method = WebRequestMethods.Http.Get;
                        webRequestTR.Host = "graph.microsoft.com";
                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                        webRequestTR.AllowAutoRedirect = false;

                        HttpWebResponse resp = (HttpWebResponse)webRequestTR.GetResponse();
                        string redirUrl = resp.Headers["Location"];
                        resp.Close();
                        resp.Dispose();

                        WebClient client = new WebClient();
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                        ServicePointManager.DefaultConnectionLimit = 9999;
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        Stream data = client.OpenRead(redirUrl);
                        StreamReader reader = new StreamReader(data);
                        string sResultCSV = reader.ReadToEnd();
                        data.Close();
                        reader.Close();

                        string[] sResultCSVArray = sResultCSV.Split('\n');
                        int iCount = 0;
                        sResult = "<table cellspacing='3' cellpadding='3'>";
                        foreach (string sCSVLine in sResultCSVArray)
                        {
                            if (sCSVLine != "")
                            {
                                sResult += "<tr>";

                                string[] sCurentLineArray = sCSVLine.Split(',');
                                for (int i = 0; i < sCurentLineArray.Length; i++)
                                {
                                    if (i != 5)
                                    {
                                        if (iCount == 0) sResult += "<th>"; else sResult += "<td>";
                                        sResult += sCurentLineArray[i];
                                        if (iCount == 0) sResult += "</th>"; else sResult += "</td>";
                                    }
                                }

                                sResult += "</tr>";
                                iCount++;
                            }                            
                        }
                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
            }

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get the current user's presence.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetMyPresence(String accessToken)
        {
            string endpoint = "https://graph.microsoft.com/beta/me/presence";
            String userpresence = "n/a";
            String useractivity = "n/a";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint, accessToken);
            if (response != null && response.IsSuccessStatusCode)
            {
                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                userpresence = json.GetValue("availability").ToString();
                useractivity = json.GetValue("activity").ToString();
            }
            return userpresence?.Trim() + ',' + useractivity?.Trim();
        }

        /// <summary>
        /// Get the current user's displayname from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetMyDisplayName(String accessToken)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me";
            string queryParameter = "?$select=displayName";
            String userdisplayName = "123";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint + queryParameter, accessToken);
            if (response != null && response.IsSuccessStatusCode)
            {
                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                userdisplayName = json.GetValue("displayName").ToString();                
            }
            return userdisplayName?.Trim();
        }

        /// <summary>
        /// Get the current user's displayname from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetMyMail(String accessToken)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me";
            string queryParameter = "?$select=mail";
            String userMail = "1234";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint + queryParameter, accessToken);
            if (response != null && response.IsSuccessStatusCode)
            {
                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                userMail = json.GetValue("mail").ToString();
            }
            return userMail?.Trim();
        }

        /// <summary>
        /// Get the current user's id from their profile.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetMyId(String accessToken)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me";
            string queryParameter = "?$select=id";
            String userId = "";
            HttpResponseMessage response = await ServiceHelper.SendRequest(HttpMethod.Get, endpoint + queryParameter, accessToken);
            if (response != null && response.IsSuccessStatusCode)
            {
                var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                userId = json.GetValue("id").ToString();
            }
            return userId?.Trim();
        }

        public async Task<IEnumerable<Team>> GetMyTeams(string accessToken)
        {
            return await HttpGetList<Team>($"/me/joinedTeams");
        }

        public async Task<IEnumerable<Group>> GetMyGroups(string accessToken)
        {
            return await HttpGetList<Group>($"/me/joinedGroups", endpoint: graphBetaEndpoint);
        }

        public async Task PostMessage(string accessToken, string teamId, string channelId, string message)
        {
            await HttpPost($"/teams/{teamId}/channels/{channelId}/chatThreads",
                new PostMessage()
                {
                    rootMessage = new RootMessage()
                    {
                        body = new Message()
                        {
                            content = message
                        }
                    }
                },
                endpoint: graphBetaEndpoint);
        }

        public async Task<Group> CreateNewTeamAndGroup(string accessToken, String displayName, String mailNickname, String description)
        {
            // create group
            Group groupParams = new Group()
            {
                displayName = displayName,
                mailNickname = mailNickname,
                description = description,

                groupTypes = new string[] { "Unified" },
                mailEnabled = true,
                securityEnabled = false,
                visibility = "Private",
            };

            Group createdGroup = (await HttpPost($"/groups", groupParams))
                            .Deserialize<Group>();
            string groupId = createdGroup.id;

            // add me as member
            string me = await GetMyId(accessToken);
            string payload = $"{{ '@odata.id': '{GraphRootUri}/users/{me}' }}";
            HttpResponseMessage responseRef = await ServiceHelper.SendRequest(HttpMethod.Post,
                $"{GraphRootUri}/groups/{groupId}/members/$ref",
                accessToken, payload);

            // create team
            await AddTeamToGroup(groupId, accessToken);
            return createdGroup;
        }

        public async Task AddTeamToGroup(string groupId, string accessToken)
        {
            await HttpPut($"/groups/{groupId}/team",
                new Team()
                {
                    guestSettings = new TeamGuestSettings()
                    {
                        allowCreateUpdateChannels = false,
                        allowDeleteChannels = false
                    }
                });
        }

        public async Task UpdateTeam(string teamId, string accessToken)
        {
            await HttpPatch($"/teams/{teamId}",
                new Team()
                {
                    guestSettings = new TeamGuestSettings() { allowCreateUpdateChannels = true, allowDeleteChannels = false }
                });
        }

        public async Task AddMember(string teamId, string upn, bool isOwner = false)
        {
            // If you have a user's UPN, you can add it directly to a group, but then there will be a 
            // significant delay before Microsoft Teams reflects the change. Instead, we find the user 
            // object's id, and add the ID to the group through the Graph beta endpoint, which is 
            // recognized by Microsoft Teams much more quickly. See 
            // https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/teams_api_overview 
            // for more about delays with adding members.

            // Step 1 -- Look up the user's id from their UPN
            String userId = (await HttpGet<User>($"/users/{upn}")).id;

            // Step 2 -- add that id to the group
            string payload = $"{{ '@odata.id': '{graphBetaEndpoint}/users/{userId}' }}";
            await HttpPost($"/groups/{teamId}/members/$ref", payload);

            if (isOwner)
                await HttpPost($"/groups/{teamId}/owners/$ref", payload);
        }
    }
}
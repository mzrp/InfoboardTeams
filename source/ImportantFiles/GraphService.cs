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
using System.Data.SqlClient;
using System.Configuration;
using System.Threading;
using System.Globalization;

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

    public class MsGraphSubscription
    {
        [JsonProperty(PropertyName = "@odata.context")]
        public string ODataContext { get; set; }

        public string id { get; set; }
        public string resource { get; set; }
        public string applicationId { get; set; }
        public string changeType { get; set; }
        public string clientState { get; set; }
        public string notificationUrl { get; set; }
        public DateTime expirationDateTime { get; set; }
        public string creatorId { get; set; }
        public string latestSupportedTlsVersion { get; set; }
    }

    public class User
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string tenantId { get; set; }
    }

    public class Organizer
    {
        public User user { get; set; }
    }

    public class Participant
    {
        public User user { get; set; }
    }

    public class CallRecord
    {
        [JsonProperty(PropertyName = "@odata.context")]
        public string OdataContext { get; set; }
        public int version { get; set; }
        public string type { get; set; }
        public IList<string> modalities { get; set; }
        public DateTime lastModifiedDateTime { get; set; }
        public DateTime startDateTime { get; set; }
        public DateTime endDateTime { get; set; }
        public string id { get; set; }
        public Organizer organizer { get; set; }
        public IList<Participant> participants { get; set; }
    }

    public class GraphService : HttpHelpers
    {
        private string GraphRootUri = ConfigurationManager.AppSettings["ida:GraphRootUri"];
        private string DatabaseConnectionString = ConfigurationManager.AppSettings["ida:DatabaseConnectionString"];

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

        /// <summary>
        /// Create Subscription
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> CreateGraphSubscription(string accessToken)
        {
            string sResult = "n/a";

            try
            {
                var webRequestSUB = WebRequest.Create("https://graph.microsoft.com/v1.0/subscriptions") as HttpWebRequest;
                if (webRequestSUB != null)
                {
                    webRequestSUB.Method = WebRequestMethods.Http.Post;
                    webRequestSUB.Host = "graph.microsoft.com";
                    webRequestSUB.Headers.Add("Authorization", "Bearer " + accessToken);
                    webRequestSUB.ContentType = "application/json";

                    DateTime ExpirationDateTime = DateTime.UtcNow + new TimeSpan(1, 0, 0, 0);
                    string sExpirationDateTime = ExpirationDateTime.Year.ToString().PadLeft(4, '0') + "-";
                    sExpirationDateTime += ExpirationDateTime.Month.ToString().PadLeft(2, '0') + "-";
                    sExpirationDateTime += ExpirationDateTime.Day.ToString().PadLeft(2, '0') + "T";
                    sExpirationDateTime += ExpirationDateTime.Hour.ToString().PadLeft(2, '0') + ":";
                    sExpirationDateTime += ExpirationDateTime.Minute.ToString().PadLeft(2, '0') + ":";
                    sExpirationDateTime += ExpirationDateTime.Second.ToString().PadLeft(2, '0');
                    sExpirationDateTime += ".9356913Z";

                    //sExpirationDateTime += ExpirationDateTime.Millisecond.ToString().PadLeft(3, '0');
                    //sExpirationDateTime += ".9356913Z";
                    //DateTime exDT = DateTime.UtcNow + new TimeSpan(0, 0, 15, 0);

                    string sParams = "{ ";
                    sParams += "\"changeType\": \"created\", ";
                    sParams += "\"notificationUrl\": \"https://infoboardteams.azurewebsites.net/Home/GetCallRecords\", ";
                    sParams += "\"resource\": \"/communications/callRecords\", ";
                    sParams += "\"expirationDateTime\": \"" + sExpirationDateTime + "\", ";
                    sParams += "\"clientState\": \"" + Guid.NewGuid().ToString() + "\", ";
                    sParams += "\"latestSupportedTlsVersion\": \"v1_2\" ";
                    sParams += "}";

                    // create subsccription starts
                    DatabaseService dbService = new DatabaseService();
                    string sLogEntry = sParams;
                    string sLogDate = DateTime.Now.ToString();
                    string sLogName = "MS Graph Validation";
                    string sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                    sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                    string sDbOk = dbService.InsertUpdateDatabase(sSQL);

                    var data = Encoding.ASCII.GetBytes(sParams);
                    webRequestSUB.ContentLength = data.Length;

                    using (var sW = webRequestSUB.GetRequestStream())
                    {
                        sW.Write(data, 0, data.Length);
                    }

                    using (var rW = webRequestSUB.GetResponse().GetResponseStream())
                    {
                        using (var srW = new StreamReader(rW))
                        {
                            var sExportAsJson = srW.ReadToEnd();

                            // log subscription creation
                            sLogEntry = sExportAsJson.Replace("'", "\"");
                            sLogDate = DateTime.Now.ToString();
                            sLogName = "Subscription";
                            sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                            sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                            sDbOk = dbService.InsertUpdateDatabase(sSQL);
                            sResult = "Ok";
                        }
                    }

                    webRequestSUB = null;
                }
            }
            catch (Exception ex)
            {
                // create subsccription error
                DatabaseService dbService = new DatabaseService();
                string sLogEntry = ex.ToString().Replace("'", "\"");
                string sLogDate = DateTime.Now.ToString();
                string sLogName = "MS Graph Validation";
                string sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                string sDbOk = dbService.InsertUpdateDatabase(sSQL);

                sResult = ex.ToString();
            }

            return sResult;
        }

        /// <summary>
        /// Get Auth Token
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

                            sResult = "Ok";
                        }
                    }

                    webRequestAUTH = null;
                }
            }
            catch (Exception ex)
            {
                sResult = ex.ToString();
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
        /// Get report - call reocrds details.
        /// </summary>
        /// <param name="accessToken">Access token to validate user</param>
        /// <returns></returns>
        public async Task<string> GetDashboardGraphData(String accessToken)
        {
            string sResult = "n/a";

            string sStartDate = DateTime.Now.Year.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T00:00:00";
            string sEndDate = DateTime.Now.Year.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T23:59:59";

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            string strSqlQuery = "SELECT [LogEntry] FROM [dbo].[Log]  WHERE [LogName] = 'CallRecord' AND [LogDate] > '" + sStartDate + "' AND [LogDate] < '" + sEndDate + "' ORDER BY [LogDate] DESC";

            SqlConnection DatabaseFile = new SqlConnection(@DatabaseConnectionString);
            DatabaseFile.Open();

            try
            {
                using (SqlCommand commandSqlTeams = new SqlCommand(strSqlQuery, DatabaseFile))
                {
                    using (SqlDataReader reader = commandSqlTeams.ExecuteReader())
                    {
                        sResult = "<table cellspacing='3' cellpadding='3'>";

                        sResult += "<tr>";
                        sResult += "<th>Organizer</th>";
                        sResult += "<th>Participants</th>";
                        sResult += "<th>Start</th>";
                        sResult += "<th>End</th>";
                        sResult += "</tr>";

                        bool bCallsExists = false;

                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                string sExportAsJson = reader.GetString(0);
                                CallRecord notifiedCall = JsonConvert.DeserializeObject<CallRecord>(sExportAsJson.ToString());

                                sResult += "<tr>";
                                sResult += "<td>" + notifiedCall.organizer.user.displayName + "</td>";

                                sResult += "<td>";
                                for (int i = 0; i < notifiedCall.participants.Count; i++)
                                {
                                    sResult += notifiedCall.participants[i].user.displayName;
                                    if (i < notifiedCall.participants.Count - 1)
                                    {
                                        sResult += ", ";
                                    }
                                }
                                sResult += "</td>";

                                sResult += "<td>" + notifiedCall.startDateTime.ToString() + "</td>";
                                sResult += "<td>" + notifiedCall.endDateTime.ToString() + "</td>";

                                sResult += "</tr>";

                                bCallsExists = true;
                            }
                        }

                        if (bCallsExists == false)
                        {
                            sResult += "<tr><td colspan='4'>No calls exist today (" + DateTime.Now.ToString() + ")</td></tr>";
                        }

                        sResult += "</table>";
                    }
                }
            }
            catch (Exception ex)
            {
                // log call records error
                DatabaseService dbService = new DatabaseService();
                string sLogEntry = ex.ToString().Replace("'", "\"");
                string sLogDate = DateTime.Now.ToString();
                string sLogName = "CallRecord Error";
                string sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                string ssDbOk = dbService.InsertUpdateDatabase(sSQL);
                sResult = "";
            }

            DatabaseFile.Close();

            System.Web.HttpContext.Current.Session["userDetails"] = sResult;
            return sResult;
        }

        /// <summary>
        /// Get report - device distribution details.
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
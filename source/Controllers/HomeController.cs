/* 
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
 *  See LICENSE in the source repository root for complete license information. 
 */

using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft_Teams_Graph_RESTAPIs_Connect.Auth;
using Microsoft_Teams_Graph_RESTAPIs_Connect.Models;
using Resources;
using System;

using System.Net.Http;
using Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Configuration;
using System.Collections.Generic;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.IO;

namespace GraphAPI.Web.Controllers
{
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

    public class ResourceData
    {
        public string oDataType { get; set; }
        public string oDataId { get; set; }
        public string id { get; set; }
    }

    public class Notification
    {
        public string subscriptionId { get; set; }
        public string clientState { get; set; }
        public string changeType { get; set; }
        public string resource { get; set; }
        public DateTime subscriptionExpirationDateTime { get; set; }
        public ResourceData resourceData { get; set; }
        public string tenantId { get; set; }
    }

    public class HomeController : Controller
    {
        public static bool hasAppId = ServiceHelper.AppId != "Enter AppId of your application";

        public HomeController()
        {
            // do nothing        
        }

        [HttpPost]
        public async Task<ActionResult> GetCallRecords()
        {
            // log graph event
            DatabaseService dbService = new DatabaseService();
            string sLogEntry = "";
            string sLogDate = "";
            string sLogName = "";
            string sSQL = "";
            string sDbOk = "";

            if (Request.QueryString["validationToken"] != null)
            {
                var token = Request.QueryString["validationToken"];

                // log graph validation
                sLogEntry = token;
                sLogDate = DateTime.Now.ToString();
                sLogName = "MS Graph Validation";
                sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                sDbOk = dbService.InsertUpdateDatabase(sSQL);

                return Content(token, "plain/text");
            }

            try
            {
                using (var inputStream = new System.IO.StreamReader(Request.InputStream))
                {
                    JObject jsonObject = JObject.Parse(inputStream.ReadToEnd());

                    if (jsonObject != null)
                    {
                        JArray value = JArray.Parse(jsonObject["value"].ToString());

                        foreach (var notification in value)
                        {
                            // log notification
                            sLogEntry = notification.ToString().Replace("'", "\"");
                            sLogDate = DateTime.Now.ToString();
                            sLogName = "Notification";
                            sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                            sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                            sDbOk = dbService.InsertUpdateDatabase(sSQL);

                            Notification notifiedCall = JsonConvert.DeserializeObject<Notification>(notification.ToString());
                            string sNotificationId = notifiedCall.resourceData.id;

                            try
                            {
                                string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();

                                if (accessToken != "")
                                {
                                    HttpWebRequest webRequestTR = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/communications/callRecords/" + sNotificationId) as HttpWebRequest;
                                    if (webRequestTR != null)
                                    {
                                        webRequestTR.Method = WebRequestMethods.Http.Get;
                                        webRequestTR.Host = "graph.microsoft.com";
                                        webRequestTR.Headers.Add("Authorization", "Bearer " + accessToken);
                                        webRequestTR.ContentType = "application/json";

                                        using (var rW = webRequestTR.GetResponse().GetResponseStream())
                                        {
                                            using (var srW = new StreamReader(rW))
                                            {
                                                var sExportAsJson = srW.ReadToEnd();

                                                // log call record
                                                sLogEntry = sExportAsJson.ToString().Replace("'", "\"");
                                                sLogDate = DateTime.Now.ToString();
                                                sLogName = "CallRecord";
                                                sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                                                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                                                sDbOk = dbService.InsertUpdateDatabase(sSQL);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // log call record error
                                sLogEntry = ex.ToString().Replace("'", "\"");
                                sLogDate = DateTime.Now.ToString();
                                sLogName = "CallRecord Error";
                                sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                                sDbOk = dbService.InsertUpdateDatabase(sSQL);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // log graph subscription error
                sLogEntry = ex.ToString().Replace("'", "\"");
                sLogDate = DateTime.Now.ToString();
                sLogName = "Subscription Error";
                sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                sDbOk = dbService.InsertUpdateDatabase(sSQL);
            }

            return new HttpStatusCodeResult(202);

        }

        private async Task<string> CreateGraphSubscription()
        {
            string sResult = "n/a";

            GraphService graphService = new GraphService();

            // create subscription
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            sResult = await graphService.CreateGraphSubscription(accessToken);

            return sResult;
        }

        private async Task<string> GetAuthToken()
        {
            GraphService graphService = new GraphService();
            string sResult = await graphService.GetAuthToken();
            return sResult;
        }

        private async Task<string> GetUserCountsGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsUserCounts(accessToken);

            return sResult;
        }

        private async Task<string> GetActivityCountsGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsActivityCounts(accessToken);

            return sResult;
        }

        private async Task<string> GetUserDetialsGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsUserActivity(accessToken);

            return sResult;
        }

        private async Task<string> GetDeviceUsageGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsDeviceUsage(accessToken);

            return sResult;
        }

        private async Task<string> GetDeviceCountsGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsDeviceCounts(accessToken);

            return sResult;
        }

        private async Task<string> GetDeviceDistributionGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetTeamsDeviceDistribution(accessToken);

            return sResult;
        }

        private async Task<string> GetDashboardGraphData()
        {
            GraphService graphService = new GraphService();

            // Get an access token.
            string accessToken = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
            string sResult = await graphService.GetDashboardGraphData(accessToken);

            return sResult;
        }

        private async Task<string> HandleGraphToken()
        {
            string sResult = "n/a";

            bool bGetTokenNow = false;
            try
            {
                DateTime DateJustNow = DateTime.Now;
                DateTime TokenExpriesAt = (DateTime)System.Web.HttpContext.Current.Session["AuthTokenExpireIn"];

                if (TokenExpriesAt > DateJustNow)
                {
                    bGetTokenNow = true;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                bGetTokenNow = true;
            }

            // database access
            DatabaseService dbService = new DatabaseService();

            // get token first
            if (bGetTokenNow == true)
            {
                string sTokenOk = await GetAuthToken();
                string sLogEntry = "";
                if (sTokenOk == "Ok")
                {
                    sLogEntry = System.Web.HttpContext.Current.Session["AuthToken"].ToString();
                }
                else
                {
                    sLogEntry = sTokenOk;
                }
                string sLogDate = DateTime.Now.ToString();
                string sLogName = "AuthToken";
                string sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
                sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
                string sDbOk = dbService.InsertUpdateDatabase(sSQL);
            }

            // subscribe caller notifiations

            return sResult;
        }

        private async Task<string> HandleGraphSubscription()
        {
            string sResult = "n/a";

            DatabaseService dbService = new DatabaseService();
            string sSubscription = dbService.GetLatesSubscription();

            if (sSubscription != "n/a")
            {
                string[] sSubscriptionArray = sSubscription.Split('ђ');
                MsGraphSubscription subscription = JsonConvert.DeserializeObject<MsGraphSubscription>(sSubscriptionArray[0].ToString());

                if (DateTime.UtcNow > subscription.expirationDateTime)
                {
                    try
                    {
                        sResult = await CreateGraphSubscription();
                    }
                    catch (Exception ex)
                    {
                        ex.ToString();
                    }
                }
            }

            return sResult;
        }

        [OutputCache(NoStore = true, Duration = 0)]
        public async Task<ActionResult> Ora(string TRP)
        {
            await HandleGraphToken();
            await HandleGraphSubscription();

            if (TRP == "1")
            {
                await GetUserDetialsGraphData();
            }

            if (TRP == "2")
            {
                await GetUserCountsGraphData();
            }

            if (TRP == "3")
            {
                await GetActivityCountsGraphData();
            }

            if (TRP == "4")
            {
                await GetDeviceUsageGraphData();
            }

            if (TRP == "5")
            {
                await GetDeviceCountsGraphData();
            }

            if (TRP == "6")
            {
                await GetDeviceDistributionGraphData();
            }

            if (TRP == "7")
            {
                await GetDashboardGraphData();
            }

            return PartialView();
        }

        [Authorize]
        public async Task<ActionResult> DeviceUsage()
        {
            return View("DeviceUsage");
        }

        [Authorize]
        public async Task<ActionResult> DeviceCounts()
        {
            return View("DeviceCounts");
        }

        [Authorize]
        public async Task<ActionResult> DeviceDistribution()
        {
            return View("DeviceDistribution");
        }

        [Authorize]
        public async Task<ActionResult> Dashboard()
        {
            return View("Dashboard");
        }

        [Authorize]
        public async Task<ActionResult> ActivityCounts()
        {
            return View("ActivityCounts");
        }

        [Authorize]
        public async Task<ActionResult> UserCounts()
        {
            return View("UserCounts");
        }

        [Authorize]
        public async Task<ActionResult> Index()
        {
            return View("Graph");
        }

        public ActionResult About()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Administration()
        {
            return View();
        }

    }
}
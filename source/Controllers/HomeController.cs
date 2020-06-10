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

namespace GraphAPI.Web.Controllers
{
    public class HomeController : Controller
    {
        public static bool hasAppId = ServiceHelper.AppId != "Enter AppId of your application";

        public HomeController()
        {
            // do nothing
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

        [OutputCache(NoStore = true, Duration = 0)]
        public async Task<ActionResult> Ora(string TRP)
        {
            //await GetLoggedUserGraphData();
            await GetAuthToken();

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
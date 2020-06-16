using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Data.SqlClient;
using Resources;
using System.Configuration;
using System.Threading;
using System.Globalization;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles
{
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

    public class DatabaseService
    {
        private string DatabaseConnectionString = ConfigurationManager.AppSettings["ida:DatabaseConnectionString"];

        public string GetAllNotifications()
        {
            string sResult = "";

            string sStartDate = DateTime.Now.Year.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T00:00:00";
            string sEndDate = DateTime.Now.Year.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T23:59:59";

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            string strSqlQuery = "SELECT [LogEntry] FROM [dbo].[Log]  WHERE [LogName] = 'Notification' AND [LogDate] > '" + sStartDate + "' AND [LogDate] < '" + sEndDate + "' ORDER BY [LogDate] DESC";

            SqlConnection DatabaseFile = new SqlConnection(@DatabaseConnectionString);
            DatabaseFile.Open();

            try
            {
                using (SqlCommand commandSqlTeams = new SqlCommand(strSqlQuery, DatabaseFile))
                {
                    using (SqlDataReader reader = commandSqlTeams.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                Notification notifiedCall = JsonConvert.DeserializeObject<Notification>(reader.GetString(0));
                                sResult += notifiedCall.resourceData.id + "ђ";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                sResult = "n/a";
            }

            DatabaseFile.Close();

            return sResult;
        }

        public string GetLatesSubscription()
        {
            string sResult = "n/a";

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            string strSqlQuery = "SELECT TOP 1 L.[LogEntry], L.[LogDate] FROM [dbo].[Log] L WHERE L.[LogName] = 'Subscription' ORDER BY [Id] DESC";

            SqlConnection DatabaseFile = new SqlConnection(@DatabaseConnectionString);
            DatabaseFile.Open();

            try
            {
                using (SqlCommand commandSqlTeams = new SqlCommand(strSqlQuery, DatabaseFile))
                {
                    using (SqlDataReader reader = commandSqlTeams.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if ((!reader.IsDBNull(0)) && (!reader.IsDBNull(1)))
                            {
                                sResult = reader.GetString(0) + "ђ" + reader.GetDateTime(1).ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            DatabaseFile.Close();

            return sResult;
        }

        public string InsertUpdateDatabase(string SQL)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            string sResult = "";

            SqlConnection DatabaseFile = new SqlConnection(@DatabaseConnectionString);
            DatabaseFile.Open();

            try
            {
                SqlCommand SqlCommand;
                SqlCommand = new SqlCommand(SQL, DatabaseFile);
                SqlCommand.ExecuteNonQuery();
                sResult = "DBOK";
            }
            catch (Exception ex)
            {
                sResult = "DBERROR: " + ex.ToString();
            }

            DatabaseFile.Close();

            return sResult;
        }

    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Owin;
using Owin;
using Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles;

[assembly: OwinStartup(typeof(Microsoft_Teams_Graph_RESTAPIs_Connect.Startup))]

namespace Microsoft_Teams_Graph_RESTAPIs_Connect
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);

            // log first entry
            DatabaseService dbService = new DatabaseService();
            string sLogEntry = "Infoboard Teams started.";
            string sLogDate = DateTime.Now.ToString();
            string sLogName = "Info";
            string sSQL = "INSERT INTO [dbo].[Log] ([LogEntry], [LogDate], [LogName]) ";
            sSQL += "VALUES ('" + sLogEntry + "', '" + sLogDate + "', '" + sLogName + "')";
            string sDbOk = dbService.InsertUpdateDatabase(sSQL);
        }
    }
}

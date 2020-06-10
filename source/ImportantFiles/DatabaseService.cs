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

namespace Microsoft_Teams_Graph_RESTAPIs_Connect.ImportantFiles
{
    public class DatabaseService
    {
        private string DatabaseConnectionString = ConfigurationManager.AppSettings["ida:DatabaseConnectionString"];

    }
}
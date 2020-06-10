using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Collections.Generic;
using System.Web.Mvc;

namespace Microsoft_Teams_Graph_RESTAPIs_Connect.Models
{
    public class TenantUsers 
    {
        public List<SelectListItem> userList { get; set; }
        public List<SelectListItem> selectedUser { get; set; }
    }
}
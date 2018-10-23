using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Zhaopin.ExchangeApi.Models
{
    /// <summary>
    /// 用户分组信息
    /// </summary>
    public class UserGroupResult
    {
        public  string Name { get; set; }

        public  string PrimarySmtpAddress { get; set; }

        public  string RequireSenderAuthenticationEnabled { get; set; }
   
        public  string AcceptMessagesOnlyFrom { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Zhaopin.ExchangeApi.Models
{
    /// <summary>
    /// 用户信息
    /// </summary>
    public class UserInfoResult
    {
        public  bool Success { get; set; }

        public  string Message { get; set; }

        public  string Identity { get; set; }
   
        public  string Name { get; set; }
        public string RecipientType { get; set; }

    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Zhaopin.ExchangeApi.Models
{

    /// <summary>
    /// 命令结果
    /// </summary>
    public class CommandResult
    {
        public  bool Success { get; set; }

        public  string Message { get; set; }

    }
}
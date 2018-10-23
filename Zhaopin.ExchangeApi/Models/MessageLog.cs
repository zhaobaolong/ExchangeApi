using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExchangeApi.Models
{
    /// <summary>
    /// 消息记录
    /// </summary>
    public class MessageLog
    {
        public  string Timestamp { get; set; }

        public  string EventId { get; set; }

        public  string Source { get; set; }
   
        public string MessageSubject { get; set; }

        public string ServerIp { get; set; }

        public string ServerHostname { get; set; }

    }
}
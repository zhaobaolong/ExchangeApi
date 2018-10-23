using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Routing;
using log4net.Config;

namespace Zhaopin.ExchangeApi
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            //配置日志记录器
            var fileInfo = new FileInfo(Server.MapPath("log4net.config"));
            XmlConfigurator.Configure(fileInfo);

            GlobalConfiguration.Configure(WebApiConfig.Register);
        }
    }
}

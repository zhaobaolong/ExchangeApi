using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Web.Http;
using System.Web.Http.Results;
using System.Web.Management;
using log4net;
using Newtonsoft.Json;
using Zhaopin.ExchangeApi.Models;

namespace Zhaopin.ExchangeApi.Controllers
{
    [RoutePrefix("Api/User")]
    public class UserController : ApiController
    {
        private readonly ILog _log = LogManager.GetLogger(typeof(UserController));

        /// <summary>
        /// 启用邮箱
        /// </summary>
        /// <param name="identity"></param>
        /// <param name="database"></param>
        /// <returns></returns>
        [Route("Enable")]
        public CommandResult EnableMailbox(string identity, string database)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Enable-Mailbox");
            command.AddParameter("Identity", identity);
            command.AddParameter("Database", database);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke();
                var userInfo = this.UserInfo(identity);
                if (userInfo.RecipientType == "UserMailbox")
                {
                    ret.Success = true;
                    ret.Message = "启用成功";
                }
                else
                {
                    ret.Success = false;
                    ret.Message = "启用失败";
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }

        /// <summary>
        /// 停用邮箱
        /// </summary>
        /// <param name="identity"></param>
        /// <returns></returns>
        [Route("Disable")]
        public CommandResult DisableMailbox(string identity)
        {
            var ret = new CommandResult();
            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Disable-Mailbox");
            command.AddParameter("Identity", identity);
            command.AddParameter("Confirm", false);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke();
                var userInfo = this.UserInfo(identity);
                if (userInfo.RecipientType == "User")
                {
                    ret.Success = true;
                    ret.Message = "停用成功";
                }
                else
                {
                    ret.Success = false;
                    ret.Message = "停用失败";
                }
                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        /// <summary>
        /// 查询用户信息
        /// </summary>
        /// <param name="identity"></param>
        /// <returns></returns>
        [Route("info")]
        public UserInfoResult UserInfo(string identity)
        {
            var ret = new UserInfoResult();
            ret.Identity = identity;

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("get-user");
            command.AddParameter("Identity", identity);
            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    if (results.Count > 0)
                    {
                        ret.Name = results[0].Properties["Name"].Value.ToString();
                        ret.RecipientType = results[0].Properties["RecipientType"].Value.ToString();
                        ret.Success = true;
                    }
                    else
                    {
                        ret.Success = false;
                        ret.Message = "命令无返回结果";
                    }
                }

            }
            catch (Exception ex)
            {
                ret.Success = false;
                ret.Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}";
            }
            finally
            {

                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
            return ret;
        }


        /// <summary>
        /// 查看组
        /// </summary>
        /// <returns></returns>
        [Route("GetDistributionGroup")]
        public List<string> GetDistributionGroup()
        {
            var ret = new List<string>();
            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Get-DistributionGroup");
            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    if (results.Count > 0)
                    {
                        this._log.Info("Succeed in Executing the Remote Powershell Command:");
                        this._log.Info(command.Commands[0].CommandText);

                        for (int i = 0; i < results.Count; i++)
                        {
                            foreach (PSPropertyInfo property in results[i].Properties)
                            {

                                if (property.Name == "Name")
                                {
                                    ret.Add(property.Value.ToString());
                                    break;
                                }

                            }
                        }
                    }
                    else
                    {
                        this._log.Info("No result resturned in Executing the Remote Powershell Command: ");
                        this._log.Info(command.Commands[0].CommandText);
                    }
                }
            }
            catch (Exception ex)
            {
                this._log.Error("获取组失败", ex);
            }
            finally
            {

                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
            return ret;
        }


        [Route("GetDistributionGroupMember")]
        public List<string> GetDistributionGroupMember(string groupname)
        {
            var ret = new List<string>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Get-DistributionGroupMember");
            command.AddArgument(groupname);
            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    if (results.Count > 0)
                    {
                        this._log.Info("Succeed in Executing the Remote Powershell Command:");
                        this._log.Info(command.Commands[0].CommandText);

                        for (int i = 0; i < results.Count; i++)
                        {
                            foreach (PSPropertyInfo property in results[i].Properties)
                            {

                                if (property.Name == "Name")
                                {
                                    ret.Add(property.Value.ToString());
                                    break;
                                }

                            }
                        }
                    }
                    else
                    {
                        this._log.Info("No result resturned in Executing the Remote Powershell Command: ");
                        this._log.Info(command.Commands[0].CommandText);
                    }
                }
            }
            catch (Exception ex)
            {
                this._log.Error("获取组成员失败", ex);
            }
            finally
            {

                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
            return ret;
        }


        [Route("NewDistributionGroup")]
        public CommandResult NewDistributionGroup(string groupname, string unitname)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("New-DistributionGroup");
            command.AddParameter("name", groupname);
            command.AddParameter("IgnoreNamingPolicy");
            command.AddParameter("OrganizationalUnit", unitname);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    if (results.Count > 0)
                    {
                        this._log.Info("Succeed in Executing the Remote Powershell Command:");
                        this._log.Info(command.Commands[0].CommandText);

                        foreach (PSPropertyInfo property in results[0].Properties)
                        {
                            this._log.Info(property.Name + ": " + property.Value);
                        }
                        ret.Success = true;
                        ret.Message = "添加成功";
                    }
                    else
                    {
                        this._log.Info("No result resturned in Executing the Remote Powershell Command: ");
                        this._log.Info(command.Commands[0].CommandText);
                        foreach (var p in command.Commands[0].Parameters)
                        {
                            this._log.Info(p.Name + ":" + p.Value.ToString());
                        }

                        ret.Success = false;
                        ret.Message = "命令无返回结果";
                    }
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("ADDDistributionGroupMember")]
        public CommandResult AddDistributionGroupMember(string groupname, string membername)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("ADD-DistributionGroupMember");
            command.AddArgument(groupname);
            command.AddParameter("Member", membername);


            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var membernames = this.GetDistributionGroupMember(groupname);
                    ret.Success = membernames.Contains(membername);
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }



        [Route("RemoveDistributionGroupMember")]
        public CommandResult RemoveDistributionGroupMember(string groupname, string membername)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Remove-DistributionGroupMember");
            command.AddArgument(groupname);
            command.AddParameter("Member", membername);
            command.AddParameter("Confirm", false);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var membernames = this.GetDistributionGroupMember(groupname);
                    ret.Success = !membernames.Contains(membername);
                }
                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("SetRequireSenderAuthenticationEnabled")]
        public CommandResult SetRequireSenderAuthenticationEnabled(string groupname, bool enable)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Set-DistributionGroup");
            command.AddArgument(groupname);
            command.AddParameter("RequireSenderAuthenticationEnabled", enable ? 1 : 0);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var retCheck = this.GetRequireSenderAuthenticationEnabled(groupname);
                    if (retCheck.Success && retCheck.Message.ToUpper() == enable.ToString().ToUpper())
                    {
                        ret.Success = true;
                        ret.Message = "设置属性成功";
                    }
                    else
                    {
                        ret.Success = false;
                        ret.Message = "设置属性失败";
                    }
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("GetRequireSenderAuthenticationEnabled")]
        public CommandResult GetRequireSenderAuthenticationEnabled(string groupname)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Get-DistributionGroup");
            command.AddArgument(groupname);
            command.AddCommand("Select-Object");
            command.AddParameter("Property", "RequireSenderAuthenticationEnabled");

            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null && results.Count > 0)
                {
                    ret.Success = true;
                    ret.Message = results[0].Properties["RequireSenderAuthenticationEnabled"].Value.ToString();
                }
                else
                {
                    ret.Success = false;
                    ret.Message = "获取属性失败";
                }
                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("AddGrantSendOnBehalfTo")]
        public CommandResult AddGrantSendOnBehalfTo(string groupname, string alias)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            var script = "Set-DistributionGroup " + groupname + " -GrantSendOnBehalfTo " + "@{Add=\"" + alias + "\"}";
            command.AddScript(script);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var users = this.GetGrantSendOnBehalfTo(groupname);
                    ret.Success = users.Contains(this.UserInfo(alias).Name);
                    ret.Message = "添加代理" + (ret.Success ? "成功" : "失败");
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("RemoveGrantSendOnBehalfTo")]
        public CommandResult RemoveGrantSendOnBehalfTo(string groupname, string alias)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            var script = "Set-DistributionGroup " + groupname + " -GrantSendOnBehalfTo " + "@{Remove=\"" + alias + "\"}";
            command.AddScript(script);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var users = this.GetGrantSendOnBehalfTo(groupname);
                    ret.Success = !users.Contains(this.UserInfo(alias).Name);
                    ret.Message = "移除代理" + (ret.Success ? "成功" : "失败");
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
        }

        [Route("GetGrantSendOnBehalfTo")]
        public List<string> GetGrantSendOnBehalfTo(string groupname)
        {
            var ret = new List<string>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Get-DistributionGroup");
            command.AddArgument(groupname);
            command.AddCommand("Select-Object");
            command.AddParameter("Property", "GrantSendOnBehalfTo");
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    foreach (PSObject obj in results)
                    {
                        if (obj.Properties.Any(property => property.Name == "GrantSendOnBehalfTo"))
                        {
                            var users =
                                (ArrayList)
                                    ((System.Management.Automation.PSObject)
                                        (results[0].Properties["GrantSendOnBehalfTo"].Value)).BaseObject;
                            ret.AddRange(from object user in users select user.ToString() into name select name.Contains("/") ? name.Substring(name.LastIndexOf("/", StringComparison.Ordinal) + 1) : name);
                        }
                    }
                }
                return ret;
            }
            catch (Exception ex)
            {
                return ret;
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
        }


        [Route("AddManagedBy")]
        public CommandResult AddManagedBy(string groupname, string alias)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            var script = "Set-DistributionGroup " + groupname + " -ManagedBy " + "@{Add=\"" + alias + "\"}";
            command.AddScript(script);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var users = this.GetManagedBy(groupname);
                    ret.Success = users.Contains(this.UserInfo(alias).Name);
                    ret.Message = "添加管理员" + (ret.Success ? "成功" : "失败");
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }

        }


        [Route("RemoveManagedBy")]
        public CommandResult RemoveManagedBy(string groupname, string alias)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            var script = "Set-DistributionGroup " + groupname + " -ManagedBy " + "@{Remove=\"" + alias + "\"}";
            command.AddScript(script);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    var users = this.GetManagedBy(groupname);
                    ret.Success = !users.Contains(this.UserInfo(alias).Name);
                    ret.Message = "移除管理员" + (ret.Success ? "成功" : "失败");
                }

                return ret;
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
        }

        [Route("GetManagedBy")]
        public List<string> GetManagedBy(string groupname)
        {
            var ret = new List<string>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Get-DistributionGroup");
            command.AddArgument(groupname);
            command.AddCommand("Select-Object");
            command.AddParameter("Property", "ManagedBy");
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                this._log.Info(JsonConvert.SerializeObject(results));

                if (results != null)
                {
                    if (results.Count > 0)
                    {

                        for (int i = 0; i < results.Count; i++)
                        {
                            if (results[i].Properties.Any(property => property.Name == "ManagedBy"))
                            {
                                var users =
                                    (ArrayList)
                                        ((System.Management.Automation.PSObject)
                                            (results[0].Properties["ManagedBy"].Value)).BaseObject;
                                ret.AddRange(from object user in users select user.ToString() into name select name.Contains("/") ? name.Substring(name.LastIndexOf("/", StringComparison.Ordinal) + 1) : name);
                            }
                        }
                    }
                    else
                    {
                        this._log.Info(command.Commands[0].CommandText);
                    }
                }
                return ret;
            }
            catch (Exception ex)
            {
                return ret;
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
        }


        [Route("SetDisplayName")]
        public CommandResult SetDisplayName(string identity, string displayName)
        {
            var ret = new CommandResult();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("Set-Mailbox");
            command.AddParameter("Identity", identity);
            command.AddParameter("DisplayName", displayName);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    ret.Success = true;
                    ret.Message = "设置显示名称成功";
                }
                else
                {
                    ret.Success = false;
                    ret.Message = "设置显示名称失败";
                }
            }
            catch (Exception ex)
            {
                return new CommandResult() { Success = false, Message = $"消息：{ex.Message}    堆栈:{ex.StackTrace}" };
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
            }
            return ret;
        }

        /// <summary>
        /// 查询发送日志
        /// </summary>
        /// <param name="sender">发送者</param>
        /// <param name="startTime">开始发送时间</param>
        /// <param name="endTime">截止发送时间</param>
        /// <param name="subject">主题</param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetSendLog")]
        public List<MessageLog> GetSendLog(string sender, DateTime startTime, DateTime endTime, string subject)
        {

            var ret = new List<MessageLog>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("get-transportservice");
            command.AddCommand("get-messagetrackinglog");
            command.AddParameter("Sender", sender);
            command.AddParameter("Start", startTime.ToString("yyyy/MM/dd HH:mm:ss"));
            command.AddParameter("End", endTime.ToString("yyyy/MM/dd HH:mm:ss"));
            command.AddParameter("MessageSubject", subject);
            command.AddCommand("Select-Object");
            command.AddParameter("Property", new string[] { "Timestamp", "EventId", "Source", "MessageSubject", "ServerIp", "ServerHostname" });
            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    foreach (var psobj in results)
                    {
                        var log = new MessageLog();
                        log.Timestamp = psobj.Properties["Timestamp"].Value.ToString();
                        log.EventId = psobj.Properties["EventId"].Value.ToString();
                        log.Source = psobj.Properties["Source"].Value.ToString();
                        log.MessageSubject = psobj.Properties["MessageSubject"].Value.ToString();
                        log.ServerIp = psobj.Properties["ServerIp"].Value.ToString();
                        log.ServerHostname = psobj.Properties["ServerHostname"].Value.ToString();
                        ret.Add(log);
                    }
                }

            }
            catch (Exception ex)
            {
                this._log.Error("获取发送日志失败", ex);
            }
            finally
            {
                runspace.Dispose();
                powershell.Dispose();
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reveive"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <param name="subject"></param>
        /// <returns></returns>

        [HttpGet]
        [Route("GetReceiveLog")]
        public List<MessageLog> GetReceiveLog(string reveive, DateTime startTime, DateTime endTime, string subject)
        {
            var ret = new List<MessageLog>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            command.AddCommand("get-transportservice");
            command.AddCommand("get-messagetrackinglog");
            command.AddParameter("Recipients", reveive);
            command.AddParameter("Start", startTime.ToString("yyyy/MM/dd HH:mm:ss"));
            command.AddParameter("End", endTime.ToString("yyyy/MM/dd HH:mm:ss"));
            command.AddParameter("MessageSubject", subject);
            command.AddCommand("Select-Object");
            command.AddParameter("Property", new string[] { "Timestamp", "EventId", "Source", "MessageSubject", "ServerIp", "ServerHostname" });
            powershell.Commands = command;

            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();
                if (results != null)
                {
                    foreach (var psobj in results)
                    {
                        var log = new MessageLog();
                        log.Timestamp = psobj.Properties["Timestamp"].Value.ToString();
                        log.EventId = psobj.Properties["EventId"].Value.ToString();
                        log.Source = psobj.Properties["Source"].Value.ToString();
                        log.MessageSubject = psobj.Properties["MessageSubject"].Value.ToString();
                        log.ServerIp = psobj.Properties["ServerIp"].Value.ToString();
                        log.ServerHostname = psobj.Properties["ServerHostname"].Value.ToString();
                        ret.Add(log);
                    }
                }

            }
            catch (Exception ex)
            {
                this._log.Error("获取发送日志失败", ex);
            }
            finally
            {
                runspace.Dispose();
                powershell.Dispose();
            }
            return ret;
        }

        /// <summary>
        /// 查看人员所属组
        /// </summary>
        /// <remarks>
        ///  Get-DistributionGroup -ResultSize unlimited -Filter "Members -like ""$((Get-Mailbox xiaolei.sun).DistinguishedName)""" |fl name, PrimarySmtpAddress, RequireSenderAuthenticationEnabled, AcceptMessagesOnlyFrom
        /// </remarks>
        /// <param name="alias"></param>
        /// <returns></returns>
        [Route("GetUserGroup")]
        public List<UserGroupResult> GetUserGroup(string alias)
        {
            var ret = new List<UserGroupResult>();

            var runspace = this.CreateRunspace();
            var powershell = PowerShell.Create();
            var command = new PSCommand();

            var script = "Get-DistributionGroup -ResultSize unlimited -Filter \"Members - like \"\"$((Get - Mailbox " + alias + ").DistinguishedName)\"\"\" |fl name,PrimarySmtpAddress,RequireSenderAuthenticationEnabled,AcceptMessagesOnlyFrom";
            command.AddScript(script);
            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();

                if (results != null)
                {
                    foreach (var psobj in results)
                    {
                        var log = new UserGroupResult();
                        log.Name = psobj.Properties["Name"].Value.ToString();
                        log.PrimarySmtpAddress = psobj.Properties["PrimarySmtpAddress"].Value.ToString();
                        log.RequireSenderAuthenticationEnabled = psobj.Properties["RequireSenderAuthenticationEnabled"].Value.ToString();
                        log.AcceptMessagesOnlyFrom = psobj.Properties["AcceptMessagesOnlyFrom"].Value.ToString();
                        ret.Add(log);
                    }
                }
            }
            catch (Exception ex)
            {
                this._log.Error("获取用户分组失败", ex);
            }
            finally
            {
                runspace.Dispose();
                powershell.Dispose();
            }
            return ret;
        }

        /// <summary>
        /// 创建运行空间
        /// </summary>
        /// <returns></returns>
        private Runspace CreateRunspace()
        {

            var liveIdconnectionUri = ConfigurationManager.AppSettings["uri"];
            var username = ConfigurationManager.AppSettings["adminUser"];
            var pwd = ConfigurationManager.AppSettings["adminPwd"];
            var password = new SecureString();
            foreach (var x in pwd)
            {
                password.AppendChar(x);
            }
            var credential = new PSCredential(username, password);
            var connectionInfo = new WSManConnectionInfo((new Uri(liveIdconnectionUri)), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
            return RunspaceFactory.CreateRunspace(connectionInfo);
        }
        
    }
}

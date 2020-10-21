using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.IO;
using System.Threading;
using System.Configuration;

namespace TPM_TrakReportsEngine
{
    class SendingMailUsingEWS
    {
        public static string EWSExchangeServer = string.Empty;
        public static bool SendMail(string To_List, string BCC_List, string CC_List, string Attach_File, string subject, string Body_Msg, string Server,string userId,string password)
        {
            EWSExchangeServer = ConfigurationManager.AppSettings["EWSExchangeServer"].ToString();
            int retrySendMail = 0;
            if(string.IsNullOrEmpty(Server) || string.IsNullOrEmpty(userId) || string.IsNullOrEmpty(password))
            {
                Logger.WriteErrorLog("Web Service URL/User id/Password NOT provided.");
                return false;
            }

            bool sendMail = false;
            while (retrySendMail<2 && !sendMail)
            {
                try
                {
                    //ExchangeService service = CreateConnection("http://mmt-exch.acemicromatic.com/EWS/Exchange.asmx");
                    //ExchangeService service = CreateConnection("https://mail.acemicromatic.com/EWS/Exchange.asmx");
                    string[] recipients = null;
                    string[] ccRecipients = null;
                    string[] BccRecipients = null;
                    ExchangeService service = CreateConnection(Server, userId, password);
                    service.Timeout = 240000;
                    EmailMessage message = new EmailMessage(service);
                    message.Subject = subject;
                    if (!string.IsNullOrEmpty(Attach_File) && File.Exists(Attach_File))
                    {
                        message.Attachments.AddFileAttachment(Attach_File);
                    }
                    message.Body = Body_Msg;
                    message.Body.BodyType = BodyType.HTML;
                    if (!string.IsNullOrEmpty(To_List))
                    {
                        recipients = To_List.Split(new char[] { ';', ',' });
                        foreach (string str in recipients)
                        {
                            message.ToRecipients.Add(str);
                        }
                    }

                    if (!string.IsNullOrEmpty(CC_List))
                    {
                        ccRecipients = CC_List.Split(new char[] { ';', ',' });
                        foreach (string str in ccRecipients)
                        {
                            message.CcRecipients.Add(str);
                        }
                    }

                    //if (!string.IsNullOrEmpty(BCC_List))
                    //{
                    //    BccRecipients = BCC_List.Split(new char[] { ';', ',' });
                    //    foreach (string str in BccRecipients)
                    //    {
                    //        message.BccRecipients.Add(str);
                    //    }
                    //}
                    //message.Save();
                    message.SendAndSaveCopy();
                    sendMail = true;
                    message.Attachments.Clear();
                }
                catch (Exception ex)
                {
                    retrySendMail = retrySendMail + 1;
                    sendMail = false;
                    Logger.WriteErrorLog(ex);
                    Thread.Sleep(3000);                  
                }            
            }
            return sendMail;

        }

        private static ExchangeService CreateConnection(String url,string userId,string password)
        {
            ExchangeService service = null;
            if (!string.IsNullOrEmpty(EWSExchangeServer))
            {
                switch (EWSExchangeServer)
                {
                    case "2007 SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "2007 SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "2007 SP3":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "2010":
                        service = new ExchangeService(ExchangeVersion.Exchange2010);
                        break;

                    case "2010 SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;

                    case "2010 SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                        break;

                    case "2013":
                        service = new ExchangeService(ExchangeVersion.Exchange2013);
                        break;
                    default:
                        service = new ExchangeService();
                        break;

                };
            }
            else
            {
                service = new ExchangeService();
            }
            try
            {
                // Hook up the cert callback to prevent error if Microsoft.NET doesn't trust the server
                ServicePointManager.ServerCertificateValidationCallback =
                delegate(
                    Object obj,
                    X509Certificate certificate,
                    X509Chain chain,
                    SslPolicyErrors errors)
                {
                    return true;
                };

                service.Url = new Uri(url);
                service.Credentials = new WebCredentials(userId, password);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Problem in connecting to EWS service. Message = " + ex.ToString());
            }
            return service;
        }

    }
}

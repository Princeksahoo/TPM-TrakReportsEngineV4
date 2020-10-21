using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using System.Reflection; //For Sending using Microsoft Outlook
using System.Web;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Net;

namespace TPM_TrakReportsEngine
{
    class SendEmail
    {
        public static string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static string ProxyIPAddress = ConfigurationManager.AppSettings["ProxyIPAddress"];
        private static int ProxyPortNo = ConfigurationManager.AppSettings["ProxyPortNo"].Equals("") ? 0 : int.Parse(ConfigurationManager.AppSettings["ProxyPortNo"]);
        private static string ProxyUsername = ConfigurationManager.AppSettings["ProxyUsername"];
        private static string ProxyPassword = ConfigurationManager.AppSettings["ProxyPassword"];

        public static void SendEmailMsg(bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string Attach_File, string FileName)
        {
            if (!Email_Flag)
            {
                return;
            }
            string emailMethod = "smtp";
            string serverName = string.Empty;
            string body_Msg = "Please Find the attachment. <br> This is an automated email from TPM-Trak. <br> Please do not reply.";
            string subject = FileName + " - Report From TPM-Trak";
            int portNo = 25;
            string userID = string.Empty;
            string password = string.Empty;
            bool Msg_send = false;            

            SqlDataReader reader = null;
            try
            {
                reader = AccessReportData.GetSendEmail();
                if (reader.Read())
                {
                    emailMethod = Convert.ToString(reader["ValueinText"]);
                    if (!Convert.IsDBNull(reader["ValueinText2"]))
                    {
                        serverName = reader["ValueinText2"].ToString();
                    }
                    if (!Convert.IsDBNull(reader["Valueinint"]))
                    {
                        portNo = int.Parse(reader["Valueinint"].ToString());
                    }
                }
                else
                {
                    Logger.WriteDebugLog("Method of email(Parameter ='ScheduledReports_Email') not set in ShopDefaults.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
                return;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }

            //Get subject/body_Msg
            SqlDataReader sdrMailSubjectBody = null;
            try
            {
                sdrMailSubjectBody = AccessReportData.GetMailSubjectAndBody();
                if (sdrMailSubjectBody.HasRows)
                {
                    if (sdrMailSubjectBody.Read())
                    {
                        subject = FileName + " " + sdrMailSubjectBody["ValueInText"].ToString();
                        body_Msg = sdrMailSubjectBody["ValueInText2"].ToString();
                    }
                }
                else
                {
                    Logger.WriteDebugLog("Email:-  subject/body_Msg NOT set for parameter 'ScheduledReportsEmail_Text' in ShopDefaults.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (sdrMailSubjectBody != null)
                {
                    sdrMailSubjectBody.Close();
                }
            }

            //Get userID/password
            if (emailMethod.Equals("smtp", StringComparison.OrdinalIgnoreCase) ||
                emailMethod.Equals("ews", StringComparison.OrdinalIgnoreCase))                
            {
                SqlDataReader sdrUserIdPassword = null;
                try
                {
                    sdrUserIdPassword = AccessReportData.GetMailServerDomain();
                    if (sdrUserIdPassword.HasRows)
                    {
                        if (sdrUserIdPassword.Read())
                        {
                            userID = sdrUserIdPassword["ValueInText"].ToString();
                            password = sdrUserIdPassword["ValueInText2"].ToString();
                        }
                    }
                    else
                    {
                        Logger.WriteDebugLog("Email:- userID/password NOT set for parameter 'ScheduledReports_MailServerDomain' in ShopDefaults.");
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex);
                    return;
                }
                finally
                {
                    if (sdrUserIdPassword != null)
                    {
                        sdrUserIdPassword.Close();
                    }
                }
            }

            switch (emailMethod.ToLower())
            {                
                case "smtp": //g:
                    Msg_send = SendEmailSMTP_MS_NET(Email_List_To, Email_List_BCC, Email_List_CC, Attach_File, subject, body_Msg, serverName, Convert.ToString(portNo), userID, password);
                    break;
                case "ews":
                    Msg_send = SendingMailUsingEWS.SendMail(Email_List_To, Email_List_BCC, Email_List_CC, Attach_File, subject, body_Msg, serverName, userID, password);
                    break;
                case "telnet":
                    Logger.WriteDebugLog("Telnet not implemented. Please use SMTP OR EWS.");
                    break;
            }

            if (emailMethod.Equals("ews", StringComparison.OrdinalIgnoreCase)) // g: verify that Http method works
            {
                try
                {
                    //string urlAddress = serverName;
                    string urlAddress = "https://google.com/";
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
                    request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0";
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        Logger.WriteDebugLog("Response: OK");
                    }
                    else
                    {
                        Logger.WriteDebugLog("Response: " + response.StatusCode.ToString());
                    }
                    response.Close();
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
            }

            if (Msg_send)
            {
                Logger.WriteDebugLog(string.Format("Email:- Mail Sent Successfully Through {0}.", emailMethod.ToUpper()));
            }
            else
            {
                Logger.WriteDebugLog(string.Format("Email:- Mail not Sent Successfully Through {0} Method.", emailMethod.ToUpper()));
            }
                                
        }

        public static bool SendEmailTelnet(string From_List, string To_List, string BCC_List, string CC_List, string Attach_File, string subject, string Body_Msg, string Server, string portNo)
        {
            try
            {
                string ExePath = Path.Combine(APath, "blat.exe ");
                To_List = To_List.Replace(";", ",");
                BCC_List = BCC_List.Replace(";", ",");
                CC_List = CC_List.Replace(";", ",");

                string PPath = string.Empty;

                if (Attach_File.Trim() == string.Empty)
                {
                    PPath = string.Format(@" -t {0} -f {1} -s ""{2}"" -body ""{3}"" -server {4} -port {5}", To_List, From_List, subject, Body_Msg, Server, portNo);
                }
                else
                {
                    PPath = string.Format(@" -t {0} -f {1} -s ""{2}"" -body ""{3}"" -server {4} -port {5} -mime -base64 -attach ""{6}""", To_List, From_List, subject, Body_Msg, Server, portNo, Attach_File);
                }
                if (File.Exists(ExePath))
                {
                    ProcessStartInfo PInfo = new ProcessStartInfo(ExePath, PPath);
                    PInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    PInfo.CreateNoWindow = false;
                    Process.Start(PInfo);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Email:- {0} Method Error:{1}.", "TELNET", ex.ToString()));
                return false;
            }
        }

        public static bool SendEmailSMTP_MS_NET(string To_List, string BCC_List, string CC_List, string Attach_File, string subject, string Body_Msg, string Server, string portNo, string userId, string password)
        {
            Attachment att = null;
            MailMessage Mail = null;
            bool sendMail = false;
            int retrySendMail = 0;

            if (!(string.IsNullOrEmpty(ProxyIPAddress) || ProxyPortNo == 0))
            {
                WebRequest.DefaultWebProxy = new WebProxy(ProxyIPAddress, ProxyPortNo);
                if (!(string.IsNullOrEmpty(ProxyUsername) || string.IsNullOrEmpty(ProxyPassword)))
                    WebRequest.DefaultWebProxy.Credentials = new NetworkCredential(ProxyUsername, ProxyPassword);
            }

            while (retrySendMail < 2 && !sendMail)
            {
                try
                {
                    string[] ccRecipients = null;
                    string[] BccRecipients = null;
                    ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;
                    SmtpClient client = new SmtpClient(Server, int.Parse(portNo));
                    client.UseDefaultCredentials = false;
                    if (!string.IsNullOrEmpty(userId) && !string.IsNullOrEmpty(password))
                    {
                        client.Credentials = new System.Net.NetworkCredential(userId, password);
                    }
                    bool enableSSL = false;
                    try
                    {
                        bool.TryParse(ConfigurationManager.AppSettings["EnableSSL"].ToString(), out enableSSL);
                    }
                    catch (Exception exx)
                    {
                        Logger.WriteErrorLog(exx.ToString());
                    } 

					client.EnableSsl = enableSSL;
                    To_List = To_List.Replace(';', ',');
                    To_List = To_List.TrimEnd(',');
                    Mail = new MailMessage(userId, To_List);
                    Mail.Subject = subject;
                    AlternateView htmlView = AlternateView.CreateAlternateViewFromString(Body_Msg, null, "text/html");
                    Mail.AlternateViews.Add(htmlView);

                    //if (!string.IsNullOrEmpty(BCC_List))
                    //{
                    //    BccRecipients = BCC_List.Split(';');
                    //    foreach (string str in BccRecipients)
                    //    {
                    //        Mail.Bcc.Add(new MailAddress(str));
                    //    }
                    //}

                    if (!string.IsNullOrEmpty(CC_List))
                    {
                        ccRecipients = CC_List.Split(new char[]{';',','});
                        foreach (string str in ccRecipients)
                        {
                            Mail.CC.Add(new MailAddress(str));
                        }
                    }

                    if (Attach_File.Trim() != "")
                    {
                        att = new Attachment(Attach_File);
                        Mail.Attachments.Add(att);
                    }
                    //200 sec timeout
                    client.Timeout = 200000;
                    client.Send(Mail);
                    Logger.WriteDebugLog("email was sent successfully using SendEmailSMTP_MS_NET method");
                    sendMail = true;
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(string.Format("Email:- {0} Method Error:{1}.", "SMTP", ex.ToString()));
                    sendMail = false;
                    if (ex.Message.ToLower().Contains("please do get messages first to authenticate yourself"))
                    {
                        authRediffPro auth = new authRediffPro();
                        string popServer = Server.ToLower().Replace("smtp", "pop");
                        int popPort = 110;
                        auth.AuthRediffProSever(popServer, popPort, userId, password);
                    }
                }
                finally
                {
                    retrySendMail = retrySendMail + 1;
                    if (Mail != null) Mail.Attachments.Clear();
                    if (att != null) att.Dispose();
                    if (Mail != null) Mail.Dispose();
                }
            }
            return sendMail;
        }
       
        
        private static void OldCode()
        {
            //SmtpClient client = new SmtpClient(Server, int.Parse(portNo));
            //client.UseDefaultCredentials = false;
            //if (!string.IsNullOrEmpty(userId) && !string.IsNullOrEmpty(password))
            //{
            //    client.Credentials = new System.Net.NetworkCredential(userId, password);
            //}
            //bool enableSSL = false;
            //try
            //{
            //    bool.TryParse(ConfigurationManager.AppSettings["EnableSSL"].ToString(), out enableSSL);
            //}
            //catch (Exception exx)
            //{
            //    Logger.WriteErrorLog(exx.ToString());
            //}

            //client.EnableSsl = enableSSL;
            //To_List = To_List.Replace(';', ',');
            //Mail = new MailMessage(userId, To_List);
            //Mail.Subject = subject;
            //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(Body_Msg, null, "text/html");
            //Mail.AlternateViews.Add(htmlView);

            //if (!string.IsNullOrEmpty(BCC_List))
            //{
            //    BccRecipients = BCC_List.Split(';');
            //    foreach (string str in BccRecipients)
            //    {
            //        Mail.Bcc.Add(new MailAddress(str));
            //    }
            //}
            //if (!string.IsNullOrEmpty(CC_List))
            //{
            //    ccRecipients = CC_List.Split(';');
            //    foreach (string str in ccRecipients)
            //    {
            //        Mail.CC.Add(new MailAddress(str));
            //    }
            //}

            //if (Attach_File.Trim() != "")
            //{
            //    att = new Attachment(Attach_File);
            //    Mail.Attachments.Add(att);
            //}
            ////200 sec timeout
            //client.Timeout = 200000;
            //client.Send(Mail);
            //sendMail = true;
        }
      

    }
}

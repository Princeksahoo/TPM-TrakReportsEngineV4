using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;
using System.IO;

namespace TPM_TrakReportsEngine
{
    class authRediffPro
    {
        public TcpClient Server;
        public NetworkStream NetStrm;
        public StreamReader RdStrm;
        public string Data;
        public byte[] szData;
        public string CRLF = "\r\n";

        public void AuthRediffProSever(string address,int portNo,string userID, string password)
        {            
            // create server POP3 with port 110
            Server = new TcpClient(address, portNo);
            StringBuilder status = new StringBuilder();

            try
            {
                // initialization
                NetStrm = Server.GetStream();
                RdStrm = new StreamReader(Server.GetStream());
                status.Append(RdStrm.ReadLine());

                // Login Process
                Data = "USER " + userID + CRLF;
                szData = System.Text.Encoding.ASCII.GetBytes(Data.ToCharArray());
                NetStrm.Write(szData, 0, szData.Length);
                status.Append(Data + RdStrm.ReadLine());

                Data = "PASS " + password + CRLF;
                szData = System.Text.Encoding.ASCII.GetBytes(Data.ToCharArray());
                NetStrm.Write(szData, 0, szData.Length);
                status.Append("PASS :" + RdStrm.ReadLine());

                // Send STAT command to get information ie: number of mail and size
                Data = "STAT" + CRLF;
                szData = System.Text.Encoding.ASCII.GetBytes(Data.ToCharArray());
                NetStrm.Write(szData, 0, szData.Length);
                status.Append(Data + RdStrm.ReadLine());


                // Send STAT command to get information ie: number of mail and size
                Data = "LIST" + CRLF;
                szData = System.Text.Encoding.ASCII.GetBytes(Data.ToCharArray());
                NetStrm.Write(szData, 0, szData.Length);
                string szTemp = RdStrm.ReadLine();
                status.Append(Data + szTemp);
                if (szTemp != null && szTemp.Contains("OK"))
                {
                    while (szTemp != ".")
                    {
                        //status.Append(szTemp + CRLF);
                        szTemp = RdStrm.ReadLine();
                    }
                }

                // Send QUIT command to close session from POP server
                Data = "QUIT" + CRLF;
                szData = System.Text.Encoding.ASCII.GetBytes(Data.ToCharArray());
                NetStrm.Write(szData, 0, szData.Length);
                status.Append(Data + RdStrm.ReadLine());
            }
            catch (InvalidOperationException err)
            {
                status.Append("Error: " + err.ToString());
                Logger.WriteErrorLog(err.ToString());
            }
            catch (Exception exp)
            {
                status.Append("Error: " + exp.ToString());
                Logger.WriteErrorLog(exp.ToString());
            }
            finally
            {                
               if(NetStrm != null) NetStrm.Close();
               if (RdStrm != null) RdStrm.Close();
               Logger.WriteDebugLog(status.ToString());
            }
        }
    }
}

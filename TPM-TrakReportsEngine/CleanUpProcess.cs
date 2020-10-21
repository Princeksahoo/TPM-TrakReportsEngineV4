using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Configuration;

namespace TPM_TrakReportsEngine
{
    class CleanUpProcess
    {
        public static void DeleteFiles(string Folder)
        {
            try
            {
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                DirectoryInfo di = new DirectoryInfo(Path.Combine(APath , Folder));
                FileInfo[] files = di.GetFiles("*.txt");

                if (files.Length > 0)
                {
                    int Adays = int.Parse(ConfigurationManager.AppSettings["LogHistoryDays"].ToString());
                    foreach (FileInfo fi in files)
                    {                        
                        DateTime dt1, dt2;
                        dt1 = DateTime.Now;
                        dt2 = DateTime.Parse(fi.LastWriteTime.ToString());
                        TimeSpan ts = dt1 - dt2;
                        int days = ts.Days + 1;
                       
                        if (days >= Adays)
                        {
                            try
                            {
                                Logger.WriteDebugLog("Deleting the file " + fi.Name);
                                fi.Delete();
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteErrorLog(ex.Message);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               Logger.WriteErrorLog(ex);
            }
        }
    }
}

using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

namespace TPM_TrakReportsEngine
{
    public static class Logger
    {
        private static string appPath;
        private static object _locker;
        static Logger()
        {
            string ProgTime = String.Format("{0:yyyyMMdd}", DateTime.Now);
            string logPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); //Environment.CurrentDirectory;
            Directory.CreateDirectory(Path.Combine(logPath, "Logs"));
            _locker = new object();
            appPath = Path.Combine(logPath, "Logs", "ScheduleReportLog" + ProgTime + ".txt");           
        }

        public static void WriteDebugLog(string str)
        {
            StreamWriter writer = null;           
            if (Monitor.TryEnter(_locker, 1000))
            {
                try
                {
                    writer = new StreamWriter(appPath, true, Encoding.UTF8, 8195);
                    writer.WriteLine(string.Format("{0} : Debug - {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"), str));
                    writer.Flush();
                    writer.Close();
                    writer.Dispose();
                    writer = null;
                }
                catch { }
                finally
                {
                    Monitor.Exit(_locker);
                    if (writer != null)
                    {
                        writer.Close();
                        writer.Dispose();
                    }                   
                }
            }
        }

        public static void WriteErrorLog(string str)
        {
            StreamWriter writer = null;            
            if (Monitor.TryEnter(_locker, 1000))
            {
                try
                {
                    writer = new StreamWriter(appPath, true, Encoding.UTF8, 8195);
                    writer.WriteLine(string.Format("{0} : Exception - {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"), str));
                    writer.Flush();
                    writer.Close();
                    writer.Dispose();
                    writer = null;
                }
                catch { }
                finally
                {
                    Monitor.Exit(_locker);
                    if (writer != null)
                    {
                        writer.Close();
                        writer.Dispose();
                    }                    
                }
            }
        }

        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter writer = null;           
            //if (Monitor.TryEnter(_locker, 1000))
            {
                try
                {
                    writer = new StreamWriter(appPath, true, Encoding.UTF8, 8195);
                    writer.WriteLine(string.Format("{0} : Exception - {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"), ex));
                    writer.Flush();
                    writer.Close();
                    writer.Dispose();
                    writer = null;
                }
                catch { }
                finally
                {
                    //  Monitor.Exit(_locker);
                    if (writer != null)
                    {
                        writer.Close();
                        writer.Dispose();
                    }
                }
            }
        } 
    }
}

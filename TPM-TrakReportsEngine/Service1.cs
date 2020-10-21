using System;
using System.Data;
using System.Security.Permissions;
using System.ServiceProcess;
using System.Text;
using System.Threading;


namespace TPM_TrakReportsEngine
{
    public partial class TPMTrakReportsEngine : ServiceBase
    {
        Thread tr = null;
        
        public TPMTrakReportsEngine()
        {
            InitializeComponent();
        }

        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.ControlAppDomain)]
        protected override void OnStart(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            try
            {
                Logger.WriteDebugLog("Starting Service.");
                CreateClient C1 = new CreateClient();
                ThreadStart job = new ThreadStart(C1.GetClient);
                tr = new Thread(job);
                tr.Name = "ScheduledReport";
                tr.Start();
                Logger.WriteDebugLog("Service thread has been started.");
            }
            catch (Exception e)
            {
                Logger.WriteErrorLog(e.ToString());
            }                     
        }
        internal void StartDebug()
        {
            Logger.WriteDebugLog("Service started in DEBUG mode.");
            OnStart(null);
        }

        protected override void OnStop()
        {
            Logger.WriteDebugLog("Service has been stopped.");
            if (tr != null && tr.ThreadState == ThreadState.Running)
            {
                try
                {
                    tr.Abort();
                }
                catch(Exception ex)
                { }
            }
        }

        void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = args.ExceptionObject as Exception;
            if (e != null)
            {
                Logger.WriteErrorLog("Unhandled Exception caught : " + e.ToString());
                Logger.WriteErrorLog("Runtime terminating:" + args.IsTerminating);
                var threadName = Thread.CurrentThread.Name;
                Logger.WriteErrorLog("Exception from Thread = " + threadName);
                System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
                StringBuilder str = new StringBuilder();
                if (p != null)
                {
                    str.AppendLine("Total Handle count = " + p.HandleCount);
                    str.AppendLine("Total Threads count = " + p.Threads.Count);
                    str.AppendLine("Total Physical memory usage: " + p.WorkingSet64);

                    str.AppendLine("Peak physical memory usage of the process: " + p.PeakWorkingSet64);
                    str.AppendLine("Peak paged memory usage of the process: " + p.PeakPagedMemorySize64);
                    str.AppendLine("Peak virtual memory usage of the process: " + p.PeakVirtualMemorySize64);
                    Logger.WriteErrorLog(str.ToString());
                }
                Thread.CurrentThread.Abort();
                //while (true)
                //    Thread.Sleep(TimeSpan.FromHours(1));

            }
        }

    }

}

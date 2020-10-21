using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Configuration;

namespace TPM_TrakReportsEngine
{
    class CreateClient
    {
        public string CompanyName = string.Empty;
        public string shift = string.Empty;
        string _appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public CreateClient()
        {
        }
        public void GetClient()
        {
            string timeDelay = ConfigurationManager.AppSettings["TimeDelayAfterShiftEnd"];
            int logHistoryDays = int.Parse(ConfigurationManager.AppSettings["LogHistoryDays"].ToString());
            int timeDelayAfterShiftEnd = 30;
            if (!string.IsNullOrEmpty(timeDelay))
            {
                timeDelayAfterShiftEnd = Convert.ToInt32(timeDelay);
            }
            DateTime CDT = DateTime.Now.Date.AddDays(1);
            CompanyName = AccessReportData.GetCompanyName();
            GenerateReport();

            DateTime curShiftEndTime = GetCurrentShiftEndTime();
            curShiftEndTime = curShiftEndTime.AddMinutes(timeDelayAfterShiftEnd);
            Logger.WriteDebugLog("Reports will be exported at " + curShiftEndTime.ToString("yyyy-MM-dd HH:mm:ss"));
            while (true)
            {
                try
                {
                    ExportOnDemandReports();
                    if (CDT < DateTime.Now)
                    {
                        CDT = CDT.AddDays(logHistoryDays);
                        CleanUpProcess.DeleteFiles("Logs");
                    }
                    if (curShiftEndTime < DateTime.Now)
                    {
                        GenerateReport();
                        curShiftEndTime = GetCurrentShiftEndTime();
                        curShiftEndTime = curShiftEndTime.AddMinutes(timeDelayAfterShiftEnd);
                        Logger.WriteDebugLog("Next time reports will be exported at " + curShiftEndTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    Thread.Sleep(10000);
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                    Thread.Sleep(10000);
                }
            }
        }

        public void GetStartEnd(string Vartime, string ScheduledReports_LastRunforTheDay, out string StartTime, out string EndTime, out string shiftlog)
        {
            EndTime = string.Empty;
            StartTime = string.Empty;
            shiftlog = string.Empty;
            bool IsPDT = AccessReportData.GetPDT(ScheduledReports_LastRunforTheDay);
            if (Vartime.ToLower() == "day")
            {
                if (IsPDT)
                {
                    Logger.WriteDebugLog("Report not exported for the Day " + string.Format("{0:yyyy-MMM-dd}", ScheduledReports_LastRunforTheDay) + ". (PDT)");
                    return;
                }
                else
                {
                    StartTime = AccessReportData.GetLogicalDayStart(ScheduledReports_LastRunforTheDay);
                    // EndTime = DateTime.Parse(StartTime).AddDays(1).ToString("yyyy-MMM-dd hh:mm:ss tt");
                    EndTime = AccessReportData.GetLogicalDayEnd(StartTime);
                    return;
                }
            }
            else if (Vartime.ToLower() == "shift")
            {
                if (IsPDT)
                {
                    Logger.WriteDebugLog(DateTime.Now.ToString() + " Report not exported for the Shift " + string.Format("{0:yyyy-MMM-dd}", ScheduledReports_LastRunforTheDay) + ". (PDT)");
                    return;
                }
                else
                {
                    SqlDataReader DR = AccessReportData.GetPreviousShiftEndTime();
                    if (DR.Read())
                    {
                        //DR0331::geeta added from here
                        //StartTime = Convert.ToString(DR["Starttime"]);
                        //EndTime = Convert.ToString(DR["Endtime"]);
                        StartTime = DateTime.Parse(Convert.ToString(DR["Starttime"])).ToString("yyyy-MM-dd HH:mm:ss");
                        EndTime = DateTime.Parse(Convert.ToString(DR["Endtime"])).ToString("yyyy-MM-dd HH:mm:ss");
                        shiftlog = Convert.ToString(DR["Shiftname"]);

                    }
                    else
                    {
                        StartTime = string.Empty;
                        EndTime = string.Empty;
                    }
                    if (DR != null)
                    {
                        DR.Close();
                    }
                }
            }
            else if (Vartime.ToLower() == "month") // g:
            {
                StartTime = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now, "start");
                EndTime = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now, "end");
            }
        }

        public void ExportOnDemandReports()
        {
            SqlDataReader DR = null;
            try
            {
                string StartTime = string.Empty;
                string EndTime = string.Empty;
                string plnt = string.Empty;
                string machineId = string.Empty;
                string shift = string.Empty;
                string cellID = string.Empty;
                DR = AccessReportData.GetExportReports("Now");
                while (DR.Read())
                {
                    #region Report generation switch case
                    Logger.WriteDebugLog("Generating report - " + Convert.ToString(DR["ReportName"]) + " - " + Convert.ToString(DR["ReportFileName"]));
                    switch (Convert.ToString(DR["ReportName"]))
                    {
                        case "MachineWise Production Report - Format3":
                            StartTime = Convert.ToString(DR["Shift"]);//start date
                            EndTime = Convert.ToString(DR["Operator"]);//end date
                            plnt = Convert.ToString(DR["PlantID"]);
                            machineId = Convert.ToString(DR["Machine"]);

                            ExportReport.ExportDailyProductionReportTrellBorg(_appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, false, "", "");
                            #region Update ScheduledReports
                            AccessReportData.UpdateScheduleReportTrelBorg(Convert.ToString(DR["Slno"]));
                            #endregion
                            break;
                        case "JH Audit report - On Demand":
                            shift = Convert.ToString(DR["Shift"]);
                            StartTime = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now, "start");
                            plnt = Convert.ToString(DR["PlantID"]);
                            machineId = Convert.ToString(DR["Machine"]);
                            cellID = DR["GroupID"].ToString();
                            ExportReport.ExportJHAuditReport(_appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToString(DR["Machine"]),Convert.ToDateTime(StartTime), plnt, shift,cellID, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                            #region Update ScheduledReports
                            AccessReportData.UpdateScheduleReportTrelBorg(Convert.ToString(DR["Slno"]));
                            #endregion
                            break;
                        case "PM Report (Phantom Cell)":
                            StartTime = AccessReportData.GetLogicalDayStart(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            EndTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            ExportReport.ExportPMPhantomCellReport(_appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["RunReportForEvery"]));
                            #region Update ScheduledReports
                            AccessReportData.UpdateScheduleReportTrelBorg(Convert.ToString(DR["Slno"]));
                            #endregion
                            break;
                    }
                    #endregion
                }
                if (DR != null)
                {
                    DR.Close();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (DR != null && !DR.IsClosed)
                {
                    DR.Close();
                }
            }
        }

        public void ExportALLReports(string CompanyName, bool MachineAE, int overWriteFile)
        {
            SqlDataReader DR = null;
            SqlDataReader SDR = null;

            try
            {
                string StartTime = string.Empty;
                string EndTime = string.Empty;
                bool ISExcel = false;
                string APPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string Vartime = "Day";
                string plnt = string.Empty;
                int startingDayOfMonth = Convert.ToInt32(ConfigurationManager.AppSettings["MonthStartDayForProductionDetails"].ToString());
                // Delete old reports
                int days = 3;
                if (int.TryParse(ConfigurationManager.AppSettings["DeleteOldReportsDays"], out days))
                {
                    if (days > 0)
                    {
                        List<string> paths = AccessReportData.GetExportReportPaths();
                        foreach (string path in paths)
                        {
                            Utility.DeleteOldReports(path, days);
                        }
                    }
                }

                DateTime ScheduledReports_LastRunforTheDay = AccessReportData.GetLastRunforTheDay();
                TimeSpan time = DateTime.Now.Date.Subtract(ScheduledReports_LastRunforTheDay);
                string strDate = string.Empty;
                if (time.Days > 3)
                {
                    Logger.WriteDebugLog("Scheluded Report:- ScheduledReports_LastRunforTheDay is Exceeding Three Days.");
                    DateTime currentLogicalDayStartTime = DateTime.Parse(AccessReportData.GetLogicalDayStart(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                    ScheduledReports_LastRunforTheDay = currentLogicalDayStartTime.AddDays(-3);
                }

                DateTime previousShiftEndTime = GetPreviousShiftEndTime();
                while (ScheduledReports_LastRunforTheDay < previousShiftEndTime) // changedto == from < 
                {
                    string runReportForShiftDay = "Day";
                    bool isReportPresent = false;

                    DateTime lastShiftEndTime = DateTime.Parse(AccessReportData.GetLogicalDayEnd(ScheduledReports_LastRunforTheDay.ToString("yyyy-MM-dd HH:mm:ss")));
                    previousShiftEndTime = GetPreviousShiftEndTime();

                    if (lastShiftEndTime == previousShiftEndTime)
                    {
                        runReportForShiftDay = string.Empty; // both day and Shifts
                    }
                    else if (lastShiftEndTime > previousShiftEndTime)
                    {
                        runReportForShiftDay = "shift";
                        Vartime = "shift";
                    }

                    strDate = string.Format("{0:dd_MMM_yyyy}", ScheduledReports_LastRunforTheDay)+"\\";
                    string Shiftlog;
                    DR = AccessReportData.GetExportReports(runReportForShiftDay);
                    while (DR.Read())
                    {
                        isReportPresent = true;
                        string rptparam = string.Empty;
                        string Parameter = string.Empty;
                        Vartime = Convert.ToString(DR["runreportforevery"]);
                        GetStartEnd(Vartime, ScheduledReports_LastRunforTheDay.ToString("yyyy-MM-dd HH:mm:ss"), out StartTime, out EndTime, out Shiftlog);
                        int ShiftID = AccessReportData.GetShiftIDNO(Shiftlog);

						if (StartTime != string.Empty && EndTime != string.Empty)
                        {
							#region Report generation switch case
							plnt = Convert.ToString(DR["PlantID"]);
                            Logger.WriteDebugLog("Generating report - " + Vartime + " wise - " + Convert.ToString(DR["ReportFileName"]));
                            switch (Convert.ToString(DR["ReportName"]))
                            {
								//Tafe:  Plan Vs Actual Report
								case "PlanVsActualReport":
									ISExcel = true;
									ExportReport.ExportPlanVsActualReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]));
									break;

								//Tafe:  CategoryWise OEE And Loss Time Report
								case "CategoryWise OEE And Loss Time Report":
									ISExcel = true;
									ExportReport.ExportCategoryWiseOEEAndLossTimeReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["Machine"]));
									break;
                                //Tafe: Hold Report
                                case "Hold Report":
                                    ISExcel = true;
                                    ExportReport.ExportHoldReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime),Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["GroupID"]), Convert.ToString(DR["Machine"]));
                                    break;
                                //Tafe: Machine History Report
                                case "Machine History Report":
                                    ISExcel = true;
                                    ExportReport.ExportMachineHistoryReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["Machine"]));
                                    break;
                                //Tafe: Rejection Report
                                case "Rejection Report":
                                    ISExcel = true;
                                    // Rejection Report for Material
                                    ExportReport.ExportRejectionReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime),Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]), "Material");

                                    // Rejection Report for Process
                                    ExportReport.ExportRejectionReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]), "Process");
                                    break;
                                //Tafe: Batchwise Report
                                //case "Batchwise Report":
                                //    ISExcel = true;
                                //    ExportReport.ExportBatchwiseReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]), Convert.ToString(DR["ComponentID"]), "Material");
                                //    break;
                                //Tafe: Line Meter Report
                                case "Line Meter Report":
                                    ISExcel = true;
                                    ExportReport.ExportLineMeterReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["GroupID"]));
                                    break;
								//EneryMeterReport_SONA
								case "EnergyMeterReport":
									ISExcel = true;
									ExportReport.ExportEnergyMeterReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
									break;

								//Flow meter Report
								case "FlowMeterReport":
									ISExcel = true;
									ExportReport.ExportFlowMeterReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["Shift"]));
									break;
								//SONAMIS_Report
								case "SonaMISReport":
									ISExcel = true;
									ExportReport.ExportSonaMISReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["PlantID"]),Convert.ToString(DR["Shift"]));
									break;
								//baluauto
								case "Efficiency And Graph":
									ISExcel = true;
									ExportReport.ExportEfficiencyAndGraphReportMonthwise(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["Machine"]));
									break;

								//case "Efficiency And Graph":
								//	ISExcel = true;
								//	ExportReport.ExportEfficiencyAndGraphReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["Machine"]));
								//	break;
								//mangalhourly
								case "MangalHourlyChartReport":
									ISExcel = true;
									ExportReport.ExportMangalHourlychartReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["Machine"]),Convert.ToString(DR["PlantID"]));
									break;

								//mangalDowntime
								case "MangalDowntimeReport":
									ISExcel = true;
									ExportReport.ExportMangalDowntimeReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])),Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime),bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
									break;
								case "Production and Downtime Report - Daily By Hour":
                                    ISExcel = true;
                                        ExportReport.ExportDailyProdDownDaywiseExcelReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    
                                    break;                             
                                case "DNCUsageReport":
                                    ISExcel = true;
                                    ExportReport.ExportDNCUsageReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Vartime, Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    break;
                                case "Bosch_HourlyCountWith AE_Losses":
                                    ISExcel = true;
                                    ExportReport.ExportShiftProductionCountHourly(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    break;
                                case "OEE Trend":
                                    ISExcel = true;
                                    //ExportReport.ExportOEETrend(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    ExportReport.ExportOEETrend(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Shiftlog, Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    break;
                                case "Inconsistent MCO":
                                    ISExcel = true;
                                    ExportPDF.createMCOreportPDF(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, StartTime, EndTime, Convert.ToString(DR["PlantID"]), Convert.ToString(DR["Machine"]), Shiftlog, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "MachineWise Production Report - Format3":
                                    ISExcel = true;
                                    ExportReport.ExportDailyProductionReportTrellBorgDayWise(_appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, "", Convert.ToString(DR["Machine"]), "", StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, false, "", "");
                                    break;
                                // NR0114 Added By Shwetha - April 10
                                case "Cockpit":
                                    ISExcel = true;
                                    var operatorVal = Convert.ToString(DR["Operator"]);
                                    if (string.IsNullOrEmpty(operatorVal)) operatorVal = "PCT";
                                    ExportPDF.createCockpitreportPDF(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, StartTime, EndTime, Convert.ToString(DR["PlantID"]), Convert.ToString(DR["Machine"]), Shiftlog, operatorVal, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                // NR0114 Added By Shwetha - April 10
                                case "Daily OEE by Shift":
                                    ISExcel = true;
                                    ExportReport.ExportDailyProductionReportDayWiseShantiIron(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                // g:
                                case "PM Report(Shanthi)":
                                    ISExcel = true;
                                    StartTime = DateTime.Now.AddYears(-1).ToString("yyyy-MMM-dd hh:mm:ss tt");
                                    EndTime = DateTime.Now.ToString("yyyy-MMM-dd hh:mm:ss tt");
                                    ExportReport.ExportPMReportShantiIron(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Vasavi Added
                                case "Cockpit Data Report(Shanthi)":
                                    ISExcel = true;
                                    ExportReport.ExportProductionRExportCockpitProductionReportShantiIron(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Vasavi Added
                                case "Production Count By SlNo(Shanthi)":
                                    ISExcel = true;
                                    ExportReport.ExportProductionCountBySlNo(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Vasavi Added
                                case "Machinewise Alarm Report":
                                    ISExcel = true;
                                    ExportReport.ExportMachinewiseAlarmReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Vasavi Added
                                case "Monthwise OEE Report":
                                    ISExcel = true;
                                    ExportReport.ExportMonthwiseOEEReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Vasavi Added
                                case "Machine DownTime Matrix":
                                    ISExcel = true;
                                    if(!ConfigurationManager.AppSettings["AdvikPages"].ToString().Equals("1"))
                                        ExportReport.ExportMachineDownTimeMatrix(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    else
                                        ExportReport.ExportMachineDownTimeMatrix_Advik(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt,DR["GroupID"].ToString(), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "Shiftwise OEE Summary":
                                    ISExcel = true;
                                    ExportReport.ExportDailyProductionandRejectionReport(StartTime, EndTime, Shiftlog, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), ShiftID);
                                    break;
                                case "Bosch_HourlyCountWith AE_Losses_Format-I":
                                    ISExcel = true;//epplus todo vasavi
                                    //ExportReport.ExportHorlypartsCountReportFormatI(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    ExportReport.ExportShiftProductionCountHourly(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    break;
								
								//Pramod Added
								case "Sand Report(Shanthi)":
                                    ISExcel = true;
                                    ExportReport.SendFileShareFiles(bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                //Pramod Added
                                case "Float Sheet(Bosch Nashik)":
                                    ISExcel = true;
                                    ExportReport.FloatSheetGenerateReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), StartTime, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]));
                                    break;
                                case "Daily Rejection Report": //g:
                                    ISExcel = true;
                                    ExportReport.ExportDailyRejectionReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "Hourly Machinewise Production Report": //g:
                                    ISExcel = true;
                                    ExportReport.ExportHourlyMachinewiseProductionReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "Tool Life Report": //g:
                                    ISExcel = true;
                                    ExportReport.ExportToolLifeReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "EWS OEE Report": //g:
                                    ISExcel = true;
                                    ExportReport.ExportEWSOEEReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                    break;
                                case "Production and Downtimes Report": //g:
                                    ISExcel = true;
                                    ExportReport.ExportProductionAndDowntimesReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                    break;
                                case "Production and Downtime Report - Cumulative":
                                    ISExcel = true;
                                    ExportReport.ExportDailyProdDownDaywiseExcelReportOnlyLnT(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                    break;
                                //LnT - Cyclewise Production Details Report - Done by Prince
                                case "Cyclewise Production Details Report":
                                    ISExcel = true;
                                    ExportReport.ExportLnTProductionDetailsReport(Convert.ToDateTime(StartTime),Convert.ToDateTime(EndTime), _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime);
                                    break;
                                //GEA - Daily Maintenance CheckList Report - Done by Prince
                                case "Daily Maintenance CheckList Report":
                                    ISExcel = true;
                                    ExportReport.GenerateDailyChklistReport(_appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                    break;
                                case "OEE And Losstime": //g: 2019-03-18
                                    ISExcel = true;
                                    ExportReport.ExportOEEAndLosstimeReport(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), "", plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                    break;
                                //Pramod Added
                                case "Shiftwise Analysis Report(Time-Consolidated)":
                                    ISExcel = true;
                                    ExportReport.ExportProductionReportMachinewise(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), CompanyName, MachineAE);
                                    break;
                                case "JHChecklist Transaction Report":
                                    ISExcel = true;
                                    ExportReport.ExportJHChecklistTransactionReport(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]),Convert.ToString(DR["RunReportForEvery"]));
                                    break;
                                case "Machinewise Shift Production Report - Format1":
                                  ISExcel = true;
                                    ExportReport.ExportMachinewiseShiftFormat1Report(Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), "","",plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]));
                                    break;
                                case "SAP OEE Report":
                                    ExportReport.ExportSAPOEEReportAdvik(Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), plnt, DR["GroupID"].ToString(), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]),Convert.ToString(DR["RunReportForEvery"]));
                                    break;
                                case "Downtime Qualification Report":
                                    ExportReport.ExportDowntimeQualificationReportAdvik(Convert.ToDateTime(StartTime), Convert.ToDateTime(EndTime), _appPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), Convert.ToString(DR["Machine"]), plnt, DR["GroupID"].ToString(), bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), Convert.ToString(DR["RunReportForEvery"]));
                                    break;

                            }

                            switch (Convert.ToString(DR["ReportFileName"]))
                            {
                                case "SM_ShiftProductionReport.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Shiftlog, Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter); 
                                    break;
                                case "SM_ShiftWiseProdAndDown.rpt":
                                    rptparam = "ProdandDown";
                                    //ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Shiftlog, Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_PlantEffiComparisonReport.rpt":
                                    rptparam = Vartime;
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_DailyProductionReport.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_WeeklyProductionReport.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_MachineWiseProdReportfromAutodata.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_OperatorProdReportfromAutodata.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_DailyBreakDownReport.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_DailyOperatorReport.rpt":
                                    rptparam = "";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
                                case "SM_WeeklyOperatorProductionReport.rpt":
                                    Parameter = "timebatch";
                                    ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
                                    break;
								
                                default:
                                    if (!ISExcel)
                                    {
										if (DR["ReportFileName"].ToString().ToLower().Contains(".rpt"))
										{
											ExportReport.ExportCrystallReportFun(APPath + "\\Reports\\" + Convert.ToString(DR["ReportFileName"]), Convert.ToString(DR["ExportPath"]), Convert.ToString(DR["ExportFileName"]), int.Parse(Convert.ToString(DR["ExportType"])), int.Parse(Convert.ToString(DR["DayBefores"])) * -1, Convert.ToString(DR["Shift"]), Convert.ToString(DR["Machine"]), Convert.ToString(DR["operator"]), StartTime, EndTime, plnt, bool.Parse(Convert.ToString(DR["Email_Flag"])), Convert.ToString(DR["Email_List_To"]), Convert.ToString(DR["Email_List_CC"]), Convert.ToString(DR["Email_List_BCC"]), rptparam, strDate, Parameter);
										}
                                    }
                                    break;
                            }
                            ISExcel = false;
                            #endregion
                        }
                    }

                    #region toImplementWeekly
                    SDR = AccessReportData.GetExportReports("Weekly");
                    while (SDR.Read())
                    {
						Vartime = Convert.ToString(SDR["runreportforevery"]);
						//plnt = Convert.ToString(SDR["PlantID"]);
						int PreviousWeek = 0;
                        try
                        {
                            #region Machine Order Efficiency Report
                            //int PreviousWeek = 0;
                            string val = SDR["ReportName"].ToString();
                            if (Convert.ToString(SDR["ReportName"]) == "Machine Order Efficiency Report")
                            {
                                if (!Convert.IsDBNull(SDR["RunHistory"]))
                                {
                                    PreviousWeek = Utility.WeekNumber(Convert.ToDateTime(SDR["RunHistory"]));
                                }
                                if (PreviousWeek == 0 || PreviousWeek != Utility.WeekNumber(DateTime.Now))
                                {
                                    plnt = Convert.ToString(SDR["PlantID"]);
                                    //Get the previous week start and end time
                                    DateTime mondayOfLastWeek = DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6);
                                    DateTime saturdayofLastWeek = mondayOfLastWeek.AddDays(5);
                                    string weekStartDate = AccessReportData.GetLogicalDay(mondayOfLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "start");
                                    string weekMonthEndDate = AccessReportData.GetLogicalDay(saturdayofLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "end");
                                    Logger.WriteDebugLog(string.Format("Weekly MOReport generating for Week  : {0} - {1}.", mondayOfLastWeek, saturdayofLastWeek));
                                    ExportReport.ExportMOReport(APPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]), Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), int.Parse(Convert.ToString(SDR["ExportType"])), 0, "Month", Convert.ToString(SDR["Machine"]), Convert.ToString(SDR["operator"]), weekStartDate, weekMonthEndDate, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]), Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), CompanyName, MachineAE);
                                    Logger.WriteDebugLog("Weekly Report generated Successfully.");
                                    AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                                }
                            }
                            #endregion
                        }
                        catch(Exception exxx)
                        {
                            Logger.WriteErrorLog(exxx.ToString());
                        }

                        PreviousWeek = 0;
                        //to do vasavi
                        if (Convert.ToString(SDR["ReportName"]) == "Machine DownTime Matrix")
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousWeek = Utility.WeekNumber(Convert.ToDateTime(SDR["RunHistory"]));
                            }
                            if (PreviousWeek == 0 || PreviousWeek != Utility.WeekNumber(DateTime.Now))
                            {
                                plnt = Convert.ToString(SDR["PlantID"]);
                                //Get the previous week start and end time
                                DateTime mondayOfLastWeek = DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6);
                                DateTime saturdayofLastWeek = mondayOfLastWeek.AddDays(5);
                                string weekStartDate = AccessReportData.GetLogicalDay(mondayOfLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "start");
                                string weekMonthEndDate = AccessReportData.GetLogicalDay(saturdayofLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "end");
                                Logger.WriteDebugLog(string.Format("Weekly Machine DownTime Matrix Report  generating for Week  : {0} - {1}.", mondayOfLastWeek, saturdayofLastWeek));


                                ExportReport.ExportMachineDownTimeMatrix
                                (weekStartDate, weekMonthEndDate, _appPath + "\\Reports\\" +
                                Convert.ToString(SDR["ReportFileName"]),
                                Convert.ToString(SDR["ExportPath"]),
                                Convert.ToString(SDR["ExportFileName"]),
                                Convert.ToString(SDR["Machine"]), "", weekStartDate, plnt,
                                bool.Parse(Convert.ToString(SDR["Email_Flag"])),
                                Convert.ToString(SDR["Email_List_To"]),
                                Convert.ToString(SDR["Email_List_CC"]),
                                Convert.ToString(SDR["Email_List_BCC"]));
                                Logger.WriteDebugLog("Weekly Machine DownTime Matrix Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }

                        PreviousWeek = 0;
                        //to do vasavi
                        if (Convert.ToString(SDR["ReportName"]) == "Daily Production and Rejection")
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousWeek = Utility.WeekNumber(Convert.ToDateTime(SDR["RunHistory"]));
                            }
                            if (PreviousWeek == 0 || PreviousWeek != Utility.WeekNumber(DateTime.Now))
                            {
                                plnt = Convert.ToString(SDR["PlantID"]);
                                //Get the previous week start and end time
                                DateTime mondayOfLastWeek = DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6);
                                DateTime saturdayofLastWeek = mondayOfLastWeek.AddDays(5);
                                string weekStartDate = AccessReportData.GetLogicalDay(mondayOfLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "start");
                                string weekMonthEndDate = AccessReportData.GetLogicalDay(saturdayofLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "end");
                                Logger.WriteDebugLog(string.Format("Weekly Machine DownTime Matrix Report  generating for Week  : {0} - {1}.", mondayOfLastWeek, saturdayofLastWeek));


                                ExportReport.ExportDailyProductionandRejectionReport(weekStartDate, weekMonthEndDate, "", _appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]),
                                Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), "",
                                weekStartDate, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]),
                                Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), 1);


                                Logger.WriteDebugLog("Weekly Machine Production and Rejection Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }

                        PreviousWeek = 0;
                        if (SDR["ReportName"].ToString().Equals("EWS OEE Report", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousWeek = Utility.WeekNumber(Convert.ToDateTime(SDR["RunHistory"]));
                            }
                            if (PreviousWeek == 0 || PreviousWeek != Utility.WeekNumber(DateTime.Now))
                            {
                                plnt = Convert.ToString(SDR["PlantID"]);
                                //Get the previous week start and end time
                                DateTime mondayOfLastWeek = DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6);
                                DateTime saturdayofLastWeek = mondayOfLastWeek.AddDays(5);
                                string weekStartDate = AccessReportData.GetLogicalDay(mondayOfLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "start");
                                string weekMonthEndDate = AccessReportData.GetLogicalDay(saturdayofLastWeek.ToString("yyyy-MMM-dd HH:mm:ss"), "end");
                                Logger.WriteDebugLog(string.Format("Generating Weekly EWS OEE Report for Week: {0} - {1}.", mondayOfLastWeek, saturdayofLastWeek));

                                ExportReport.ExportWeeklyEWSOEEReport(weekStartDate, weekMonthEndDate, "", _appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]),
                                Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), "",
                                weekStartDate, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]),
                                Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), 1);

                                Logger.WriteDebugLog("Weekly EWS OEE Report generated Successfully");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }
                        // GEA - Weekly Maintenance Checklist Report
                        if (SDR["ReportName"].ToString().Equals("Weekly Maintenance CheckList Report", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousWeek = Utility.WeekNumber(Convert.ToDateTime(SDR["RunHistory"]));
                            }
                            if (PreviousWeek == 0 || PreviousWeek != Utility.WeekNumber(DateTime.Now))
                            {
                                ISExcel = true;
                                ExportReport.GenerateWeeklyChklistReport(_appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]), Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), Convert.ToString(SDR["PlantID"]), bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]), Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), Vartime.Equals("Day", StringComparison.OrdinalIgnoreCase));
                                Logger.WriteDebugLog("Weekly Maintenance CheckList Report generated Successfully");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }
					}

					if (SDR != null) SDR.Close();
                    #endregion
                    #region toImplementMonthWise
                    SDR = AccessReportData.GetExportReports("Month");
                    while (SDR.Read())
                    {
                        int PreviousMonth = 0;
                        plnt = Convert.ToString(SDR["PlantID"]);
                        Vartime = Convert.ToString(SDR["runreportforevery"]);

                        if (Convert.ToString(SDR["ReportName"]) == "DNCUsageReport")
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousMonth = Convert.ToDateTime(SDR["RunHistory"]).Month;
                            }
                            if (PreviousMonth == 0 || PreviousMonth != DateTime.Now.Month)
                            {
                                string monthStartDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "start");
                                string monthEndDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "end");
                                Logger.WriteDebugLog(string.Format("Monthly report generating for month : {0}.", DateTime.Now.AddMonths(-1).ToString("MMM_yyyy")));
                                ExportReport.ExportDNCUsageReport(APPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]), Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), int.Parse(Convert.ToString(SDR["ExportType"])), 0, "Month", Convert.ToString(SDR["Machine"]), Convert.ToString(SDR["operator"]), monthStartDate, monthEndDate, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]), Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), CompanyName, MachineAE);
                                Logger.WriteDebugLog("Monthly Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }

                        if (Convert.ToString(SDR["ReportName"]).Equals("PM Report(Shanthi)", StringComparison.OrdinalIgnoreCase)) // g:
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousMonth = Convert.ToDateTime(SDR["RunHistory"]).Month;
                            }
                            if (PreviousMonth == 0 || PreviousMonth != DateTime.Now.Month)
                            {
                                ISExcel = true;
                                StartTime = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddYears(-1).AddMonths(1).AddDays(-1), "start"); // e.g. if now it's march: last april to this march
                                EndTime = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddDays(-1), "end");
                                ExportReport.ExportPMReportShantiIron(StartTime, EndTime, _appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]), Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), "", StartTime, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]), Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]));
                                Logger.WriteDebugLog("PM Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }

                        #region MachineDownTime
                        if (Convert.ToString(SDR["ReportName"]) == "Machine DownTime Matrix")
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousMonth = Convert.ToDateTime(SDR["RunHistory"]).Month;
                            }
                            if (PreviousMonth == 0 || PreviousMonth != DateTime.Now.Month)
                            {
                                string monthStartDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "start");
                                string monthEndDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "end");
                                Logger.WriteDebugLog(string.Format("Monthly report generating for month : {0}.", DateTime.Now.AddMonths(-1).ToString("MMM_yyyy")));

                                ExportReport.ExportMachineDownTimeMatrix
                                (monthStartDate, monthEndDate, _appPath + "\\Reports\\" +
                                Convert.ToString(SDR["ReportFileName"]),
                                Convert.ToString(SDR["ExportPath"]),
                                Convert.ToString(SDR["ExportFileName"]),
                                Convert.ToString(SDR["Machine"]), "", monthStartDate, plnt,
                                bool.Parse(Convert.ToString(SDR["Email_Flag"])),
                                Convert.ToString(SDR["Email_List_To"]),
                                Convert.ToString(SDR["Email_List_CC"]),
                                Convert.ToString(SDR["Email_List_BCC"]));

                                Logger.WriteDebugLog("Monthly Machine DownTime Matrix Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }

                        if (Convert.ToString(SDR["ReportName"]) == "Daily Production and Rejection")
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousMonth = Convert.ToDateTime(SDR["RunHistory"]).Month;
                            }
                            if (PreviousMonth == 0 || PreviousMonth != DateTime.Now.Month)
                            {
                                string monthStartDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "start");
                                string monthEndDate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now.AddMonths(-1), "end");
                                Logger.WriteDebugLog(string.Format("Monthly report generating for month : {0}.", DateTime.Now.AddMonths(-1).ToString("MMM_yyyy")));

                                ExportReport.ExportDailyProductionandRejectionReport(monthStartDate, monthEndDate, "", _appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]),
                                Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), "",
                                monthStartDate, plnt, bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]),
                                Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), 1);

                                Logger.WriteDebugLog("Monthly Daily Production and Rejection Report generated Successfully.");
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }
                        #endregion

                        #region LnT Production Report - Monthwise
                        if(Convert.ToString(SDR["ReportName"]).Equals("Cyclewise Production Details Report", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!Convert.IsDBNull(SDR["RunHistory"]))
                            {
                                PreviousMonth = Convert.ToDateTime(SDR["RunHistory"]).Month;
                            }
                            if ((PreviousMonth == 0 || PreviousMonth != DateTime.Now.Month) && DateTime.Now.Day == startingDayOfMonth)
                            {
                                ISExcel = true;
                                DateTime startday = new DateTime(DateTime.Now.Year, DateTime.Now.Month-1, startingDayOfMonth);
                                DateTime endDay = DateTime.Now.AddDays(-1);
                                string monthStartDate = AccessReportData.GetLogicalDayStart(startday.ToString());
                                string monthEndDate = AccessReportData.GetLogicalDayEnd(endDay.ToString());
                                ExportReport.ExportLnTProductionDetailsReport(Convert.ToDateTime(monthStartDate), Convert.ToDateTime(monthEndDate), _appPath + "\\Reports\\" + Convert.ToString(SDR["ReportFileName"]), Convert.ToString(SDR["ExportPath"]), Convert.ToString(SDR["ExportFileName"]), Convert.ToString(SDR["Machine"]), bool.Parse(Convert.ToString(SDR["Email_Flag"])), Convert.ToString(SDR["Email_List_To"]), Convert.ToString(SDR["Email_List_CC"]), Convert.ToString(SDR["Email_List_BCC"]), Vartime);
                                AccessReportData.UpdateScheduleDncReportMonthWise(Convert.ToString(SDR["Slno"]));
                            }
                        }
                        #endregion

                    }
                    if (SDR != null) SDR.Close();
                    #endregion
                    if (!isReportPresent)
                    {
                        if (string.IsNullOrEmpty(Vartime))
                        {
                            Vartime = "Shift";
                        }
                        GetStartEnd(Vartime, Convert.ToString(ScheduledReports_LastRunforTheDay), out StartTime, out EndTime, out Shiftlog);
                    }
                    if (DR != null)
                    {
                        DR.Close();
                    }

                    #region Update date in ShopDefaults
                    /* To insert the LastRunfor the day setting in shop defaults, based onthe Shift and Day */
                    if (StartTime != string.Empty)
                    {
                        if (runReportForShiftDay == string.Empty || runReportForShiftDay.Equals("Day", StringComparison.OrdinalIgnoreCase))
                        {
                            EndTime = DateTime.Parse(AccessReportData.GetLogicalDayEnd(StartTime)).ToString("yyyy-MM-dd hh:mm:ss tt");
                        }
                        AccessReportData.UpdateLRunDay(EndTime);
                        ScheduledReports_LastRunforTheDay = DateTime.Parse(EndTime);
                    }
                    else
                    {
                        ScheduledReports_LastRunforTheDay = ScheduledReports_LastRunforTheDay.AddDays(1);
                        AccessReportData.UpdateLRunDay(ScheduledReports_LastRunforTheDay.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    #endregion
                }
            }
            catch (Exception ex)
			{
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (DR != null && !DR.IsClosed)
                {
                    DR.Close();
                }

                if (SDR != null && !SDR.IsClosed)
                {
                    SDR.Close();
                }
            }
        }

        private DateTime GetPreviousShiftEndTime()
        {
            string EndTime = DateTime.Now.ToString();
            string StartTime = string.Empty;
            SqlDataReader DR = AccessReportData.GetPreviousShiftEndTime();
            if (DR.Read())
            {
                //DR0331 :: Geeta added from here
                StartTime = Convert.ToString(DR["Starttime"]);
                EndTime = Convert.ToString(DR["Endtime"]);
            }
            else
            {
                StartTime = string.Empty;
                EndTime = string.Empty;
            }
            if (DR != null)
            {
                DR.Close();
            }
            return DateTime.Parse(EndTime);
        }

        public DateTime GetCurrentShiftEndTime()
        {
            SqlDataReader DR = AccessReportData.GetCurrentShiftDetails();
            DateTime EndTime = DateTime.Now;
            if (DR.HasRows)
            {
                DR.Read();
                EndTime = DateTime.Parse(Convert.ToString(DR["Endtime"]));
                if (DR != null)
                {
                    DR.Close();
                }
            }
            else
            {
                DateTime logicaldaystart = DateTime.Parse(AccessReportData.GetLogicalDayStart(DateTime.Now.ToString("yyyy-MMM-dd hh:mm:ss tt")));
                if (logicaldaystart > DateTime.Now)
                {
                    EndTime = logicaldaystart;
                }
                else
                {
                    EndTime = logicaldaystart.AddDays(1);

                }
            }
            return EndTime;
        }

        public void GenerateReport()
        {
            int overWriteFile = AccessReportData.GetOverWriteFile();
            string MacAE = AccessReportData.GetMachineAE();
            bool MachineAE = (MacAE == string.Empty) ? true : false;
            ExportALLReports(CompanyName, MachineAE, overWriteFile);
        }
    }
}

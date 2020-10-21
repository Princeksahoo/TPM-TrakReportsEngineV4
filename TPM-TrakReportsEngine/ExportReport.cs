using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Configuration;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using CrystalDecisions.Shared;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Data;
using System.Runtime.InteropServices;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Net;
using System.Data.Linq;
using TPM_TRAK_AnalyticsWebReports;
using System.util;
using System.Data;
using OfficeOpenXml.Drawing;

namespace TPM_TrakReportsEngine
{
    class ExportReport
    {
        public static string _appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public ExportReport()
        {

        }

        public static bool ExportCrystallReportFun(string strReportFile, string ExportPath, string ExportedReportFile,
            int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
            string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
            string Email_List_BCC, string rptparam, string strDate, string Parameter)
        {
            ReportDocument rpt = null;
            try
            {
                if (strReportFile.Contains("SM_DailyProductionReport.rpt") || strReportFile.Contains("SM_ShiftProductionReport.rpt"))
                {
                    ndtime = sttime;
                }
                String strDsnNane = ConfigurationManager.AppSettings["DsnName"].ToString();
                String strDatabaseName = ConfigurationManager.AppSettings["DatabaseName"].ToString();
                String strUserID = ConfigurationManager.AppSettings["UserID"].ToString();
                String strPassword = ConfigurationManager.AppSettings["Password"].ToString();
                String strwinAuth = ConfigurationManager.AppSettings["WindowsAuthentication"].ToString();

                rpt = new ReportDocument();
                //strDate = string.Format("{0:dd_MMM_yyyy}", DateTime.Now) + "\\";
                string CompanyName = AccessReportData.GetCompanyName();
                string MacAE = AccessReportData.GetMachineAE();
                bool MachineAE = (MacAE.ToLower().Equals("dont consider") == true) ? true : false;
                int FileOver = 0;
                FileOver = AccessReportData.GetOverWriteFile();
                String Rptpath, Paramname;

                Rptpath = strReportFile;
                Paramname = "";

                TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
                TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
                ConnectionInfo crConnectionInfo = new ConnectionInfo();

                Tables CrTables;

                crConnectionInfo.ServerName = strDsnNane;
                crConnectionInfo.DatabaseName = strDatabaseName;
                crConnectionInfo.UserID = strUserID;
                crConnectionInfo.Password = strPassword;
                crConnectionInfo.IntegratedSecurity = (strwinAuth.ToLower() == "true") ? true : false;


                String strProcName;
                rpt.Load(Rptpath);
                CrTables = rpt.Database.Tables;

                foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in CrTables)
                {
                    crtableLogoninfo = CrTable.LogOnInfo;
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                    Logger.WriteDebugLog(string.Format("{0} - {1} - {2} - {3}",
                        crtableLogoninfo.ConnectionInfo.ServerName,
                        crtableLogoninfo.ConnectionInfo.DatabaseName,
                        crtableLogoninfo.ConnectionInfo.UserID,
                        crtableLogoninfo.ConnectionInfo.Password
                        )
                    );
                    CrTable.ApplyLogOnInfo(crtableLogoninfo);

                    strProcName = CrTable.Location.ToString();
                    int i = strProcName.IndexOf("(");

                    if (i == -1)
                    {
                        strProcName = CrTable.Location.ToString();
                        CrTable.Location = strDatabaseName + ".dbo." + strProcName.Substring(0, strProcName.Length - 2);
                    }
                    else
                    {
                        strProcName = CrTable.Location.Substring(CrTable.Location.IndexOf("(") + 1, CrTable.Location.IndexOf(")") - (CrTable.Location.IndexOf("(")) - 3);
                        CrTable.Location = strDatabaseName + ".dbo." + strProcName;
                    }
                }
                ReportDocument subRepDoc = new ReportDocument();
                Sections crSections;
                ReportObjects crReportObjects;
                SubreportObject crSubreportObject;
                Database crDatabase;
                crSections = rpt.ReportDefinition.Sections;


                foreach (Section crSection in crSections)
                {
                    crReportObjects = crSection.ReportObjects;
                    foreach (ReportObject crReportObject in crReportObjects)
                    {
                        if (crReportObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                        {

                            //If you find a subreport, typecast the reportobject to a subreport object 
                            crSubreportObject = (SubreportObject)crReportObject;

                            //Open the subreport 
                            subRepDoc = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                            crConnectionInfo.ServerName = strDsnNane;
                            crConnectionInfo.DatabaseName = strDatabaseName;
                            crConnectionInfo.UserID = strUserID;
                            crConnectionInfo.Password = strPassword;
                            crConnectionInfo.IntegratedSecurity = (strwinAuth.ToLower() == "true") ? true : false;

                            crDatabase = subRepDoc.Database;
                            CrTables = crDatabase.Tables;
                            foreach (Table CrTable in CrTables)
                            {
                                crtableLogoninfo = CrTable.LogOnInfo;
                                crtableLogoninfo.ConnectionInfo = crConnectionInfo;

                                CrTable.ApplyLogOnInfo(crtableLogoninfo);
                                strProcName = CrTable.Location.ToString();
                                int i = strProcName.IndexOf("(");
                                if (i == -1)
                                {
                                    CrTable.Location = strDatabaseName + ".dbo." + strProcName;
                                }
                                else
                                {
                                    strProcName = CrTable.Location.Substring(CrTable.Location.IndexOf("(") + 1, CrTable.Location.IndexOf(")") - (CrTable.Location.IndexOf("(")) - 1);
                                    CrTable.Location = strDatabaseName + ".dbo." + strProcName;
                                }
                            }
                        }
                    }
                }

                foreach (ParameterField r in rpt.ParameterFields)
                {
                    Logger.WriteDebugLog(r.Name + " : " + r.PromptText + " : " + r.ReportName);
                }
                for (int i = 0; i < rpt.ParameterFields.Count; i++)
                {
                    if (rpt.ParameterFields[i].ReportName == string.Empty)
                    {

                        Paramname = rpt.ParameterFields[i].Name.ToLower();
                        Logger.WriteDebugLog(Paramname);
                        switch (rpt.ParameterFields[i].Name.ToLower())
                        {
                            case "@startdate":
                                rpt.SetParameterValue(Paramname, DateTime.Parse(sttime).AddDays(DayBefores));
                                break;
                            case "@enddate":
                                rpt.SetParameterValue(Paramname, DateTime.Parse(ndtime).AddDays(DayBefores));
                                break;
                            case "@starttime":
                                rpt.SetParameterValue(Paramname, DateTime.Parse(sttime).AddDays(DayBefores));
                                break;
                            case "@endtime":
                                rpt.SetParameterValue(Paramname, DateTime.Parse(ndtime).AddDays(DayBefores));
                                break;
                            case "@machineid":
                                rpt.SetParameterValue(Paramname, MachineId);
                                break;
                            case "@machine":
                                rpt.SetParameterValue(Paramname, MachineId);
                                break;
                            case "@componentid":
                                rpt.SetParameterValue(Paramname, "");
                                break;
                            case "@operationno":
                                rpt.SetParameterValue(Paramname, "");
                                break;
                            case "@operator":
                                rpt.SetParameterValue(Paramname, operators);
                                break;
                            case "@operatorid":
                                rpt.SetParameterValue(Paramname, operators);
                                break;
                            case "@operatorlabel":
                                rpt.SetParameterValue(Paramname, (operators == string.Empty) ? "ALL" : operators);
                                break;
                            case "@machineidlabel":
                                rpt.SetParameterValue(Paramname, (MachineId == string.Empty) ? "ALL" : MachineId);
                                break;
                            case "@shiftin":
                                rpt.SetParameterValue(Paramname, Shift);
                                break;
                            case "@plantid":
                                rpt.SetParameterValue(Paramname, plantid);
                                break;
                            case "@companyname":
                                rpt.SetParameterValue(Paramname, CompanyName);
                                break;
                            case "@downid":
                                rpt.SetParameterValue(Paramname, "");
                                break;
                            case "@parameter":
                                rpt.SetParameterValue(Paramname, Parameter);
                                break;
                            case "@exclude":
                                rpt.SetParameterValue(Paramname, 0);
                                break;
                            case "@component":
                                rpt.SetParameterValue(Paramname, "");
                                break;
                            case "@shiftname":
                                rpt.SetParameterValue(Paramname, "");
                                break;
                            case "@comparisonparam":
                                rpt.SetParameterValue(Paramname, "Shift");
                                break;
                            case "param":
                                rpt.SetParameterValue(Paramname, MachineAE);
                                break;
                            case "@param":
                                rpt.SetParameterValue(Paramname, rptparam);
                                break;
                        };
                    }
                }
                //for (int i=0;i<7;i++)
                //{
                //    Logger.WriteDebugLog(((CrystalDecisions.Shared.ParameterDiscreteValue)((new System.Collections.ArrayList(rpt.ParameterFields[i].CurrentValues)).Items[0])).Value.ToString());
                //}
                string DstFile = string.Empty;

                if (FileOver == 1)
                {
                    DstFile = Path.Combine(ExportPath, ExportedReportFile + String.Format("{0:_ddMMMyy_HHmm}{1}", DateTime.Parse(sttime).AddDays(DayBefores), (string.IsNullOrEmpty(Shift) ? string.Empty : "_" + Shift.ToUpper())));
                }
                else
                {
                    string strpath = Path.Combine(ExportPath, strDate);
                    if (!Directory.Exists(strpath))
                    {
                        Directory.CreateDirectory(strpath);
                    }
                    DstFile = Path.Combine(strpath, ExportedReportFile + String.Format("{0:_ddMMMyy_HHmm}{1}", DateTime.Parse(sttime).AddDays(DayBefores), (string.IsNullOrEmpty(Shift) ? string.Empty : "_" + Shift.ToUpper())));
                }
                string Dst = string.Empty;
                switch (ExportType)
                {
                    case 4://PDF
                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, DstFile + ".pdf");
                        Dst = DstFile + ".pdf";
                        break;
                    case 0://Excel
                        rpt.ExportToDisk(ExportFormatType.Excel, DstFile + ".xls");
                        Dst = DstFile + ".xls";
                        break;
                    case 2://Word
                        rpt.ExportToDisk(ExportFormatType.WordForWindows, DstFile + ".doc");
                        Dst = DstFile + ".doc";
                        break;
                    case 33://RTF
                        rpt.ExportToDisk(ExportFormatType.RichText, DstFile + ".rtf");
                        Dst = DstFile + ".rtf";
                        break;
                    case 1://Html
                        rpt.ExportToDisk(ExportFormatType.HTML32, DstFile + ".htm");
                        Dst = DstFile + ".htm";
                        break;
                }
                Logger.WriteDebugLog(ExportedReportFile + " Report generated sucessfully.");

                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, Dst, ExportedReportFile);
                return true;
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception while generating the report. Message = " + ex.ToString());
                return false;
            }
            finally
            {
                if (rpt != null)
                {
                    rpt.Close();
                    rpt.Dispose();
                }
            }
        }

        internal static void ExportJHAuditReport(string ReportFileName, string ExportPath, string ExportedReportFile, int ExportedType, string MachineID, DateTime startTime, string plantID, string shift, string cellID, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {
            bool isDataAvailable = false;
            
            try
            {
                string Source = string.Empty, dst = string.Empty, Template = string.Empty, time = string.Empty;
                int row = 8;

                if (!File.Exists(ReportFileName))
                {
                    Logger.WriteDebugLog("JHAuditReport Template is not found on " + ReportFileName);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("JH_Audit_Report_{0:ddMMMyyyy_HHmmss}.xlsx", DateTime.Now));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }

                File.Copy(ReportFileName, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage Excel = new ExcelPackage(newFile, true);
                var ws = Excel.Workbook.Worksheets[1];
                string Generated = string.Empty;

                int worsheetcount = 2;
                DateTime lastdate = startTime.AddMonths(1);
                int totaldays = (lastdate - startTime).Days;
                int i = 1;
                List<string> MachineIDList = new List<string>();

                var MainWorkSheet = Excel.Workbook.Worksheets[1];
                if (MachineID.Equals(string.Empty))
                {
                    MachineIDList = AccessReportData.GetMachinesbyPlantCell(plantID, cellID);
                    MachineID = "ALL";
                }                   
                else
                    MachineIDList.Add(MachineID);
                foreach (string Machine in MachineIDList)
                {
                    DataSet ChecklistDataset = AccessReportData.Getchecklistdata(Machine, shift, startTime);
                    if (ChecklistDataset != null && ChecklistDataset.Tables.Count > 0)
                    {
                        
                        for (int table = 0; table < ChecklistDataset.Tables.Count; table = table + 2)
                        {
                            int col = 1;
                            System.Data.DataTable dtshiftval = ChecklistDataset.Tables[table];
                            int Row = 8;
                            System.Data.DataTable dtoprsupval = ChecklistDataset.Tables[table + 1];
                            if (dtoprsupval != null && dtshiftval != null && dtshiftval.Rows.Count > 0 && dtshiftval.Rows.Count > 0)
                            {
                                Excel.Workbook.Worksheets.Add(Machine + " ( " + dtoprsupval.Rows[0]["ShiftName"].ToString() + " )", MainWorkSheet);
                                var workSheet = Excel.Workbook.Worksheets[worsheetcount]; worsheetcount++;
                                if (dtshiftval != null && dtshiftval.Rows.Count > 0)
                                {
                                    isDataAvailable = true;
                                    foreach (DataRow dataRow in dtshiftval.Rows)
                                    {
                                        workSheet.Cells[Row, 1].Value = col++;
                                        workSheet.Cells[Row, 2].Value = dataRow["McArea"];
                                        workSheet.Cells[Row, 3].Value = dataRow["Location"];
                                        workSheet.Cells[Row, 4].Value = dataRow["Item"];
                                        workSheet.Cells[Row, 5].Value = dataRow["CheckFor"];
                                        workSheet.Cells[Row, 6].Value = dataRow["StdCondition"];
                                        workSheet.Cells[Row, 7].Value = dataRow["CheckingMethod"];
                                        workSheet.Cells[Row, 8].Value = dataRow["AffectOnQ"];
                                        while (i <= totaldays)
                                        {
                                            workSheet.Cells[Row, (i + 8)].Value = dataRow[i.ToString()]; i++;
                                        }
                                        i = 1;
                                        Row++;
                                    }
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Top.Color.SetColor(Color.Black);
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Bottom.Color.SetColor(Color.Black);
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Right.Color.SetColor(Color.Black);
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Border.Left.Color.SetColor(Color.Black);
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[8, 1, Row - 1, 39].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C6E0B4"));
                                    workSheet.Cells[Row, 8].Value = "Operator";
                                    workSheet.Cells[Row + 1, 8].Value = "Supervisor";
                                    workSheet.Cells[Row + 2, 8].Value = "Production Head";
                                    string ProdHead = string.Empty;
                                    string ProdHeadTS = string.Empty;

                                    #region HardCoded
                                    //int adddays = 0, firstadd = 0; ;
                                    //List<string> supval = new List<string>();
                                    //DataTable newdt = new DataTable();
                                    //DayOfWeek weekstart = fromDate.DayOfWeek;
                                    //string name = string.Empty;
                                    //switch (weekstart)
                                    //{
                                    // case DayOfWeek.Monday:
                                    // firstadd = 3;
                                    // adddays = 4;
                                    // break;
                                    // case DayOfWeek.Tuesday:
                                    // firstadd = 2;
                                    // adddays = 4;
                                    // break;
                                    // case DayOfWeek.Wednesday:
                                    // firstadd = 1;
                                    // adddays = 4;
                                    // break;
                                    // case DayOfWeek.Thursday:
                                    // firstadd = 4;
                                    // adddays = 3;
                                    // break;
                                    // case DayOfWeek.Friday:
                                    // firstadd = 3;
                                    // adddays = 3;
                                    // break;
                                    // case DayOfWeek.Saturday:
                                    // firstadd = 2;
                                    // adddays = 3;
                                    // break;
                                    // case DayOfWeek.Sunday:
                                    // firstadd = 1;
                                    // adddays = 3;
                                    // break;
                                    //}
                                    //newdt = dtoprsupval.AsEnumerable().Take(firstadd).CopyToDataTable();
                                    //if ((newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Count() > 0))
                                    //{
                                    // workSheet.Cells[Row + 1, 9].Value = newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Select(x => x.Field<string>("SupervisorName")).Distinct().First().ToString() + " (" + newdt.AsEnumerable().Where(x => x.Field<DateTime?>("SupervisorTS") != null).Select(x => x.Field<DateTime?>("SupervisorTS")).Distinct().First().ToString() + " )";
                                    //}
                                    //int Col = 9; Col = Col + firstadd;
                                    //workSheet.Cells[Row + 1, 9, Row + 1, Col - 1].Merge = true;
                                    //while (Col < 39)
                                    //{
                                    // if (adddays == 3)
                                    // {
                                    // if (Col >= 39) Col = 39;
                                    // workSheet.Cells[Row + 1, Col, Row + 1, (Col + adddays - 1)].Merge = true;
                                    // newdt = dtoprsupval.AsEnumerable().Skip(firstadd).Take(adddays).CopyToDataTable();
                                    // if ((newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Count() > 0))
                                    // {
                                    // workSheet.Cells[Row + 1, Col].Value = newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Select(x => x.Field<string>("SupervisorName")).Distinct().First().ToString() + " (" + newdt.AsEnumerable().Where(x => x.Field<DateTime?>("SupervisorTS") != null).Select(x => x.Field<DateTime?>("SupervisorTS")).Distinct().First().ToString() + " )";
                                    // }
                                    // Col = Col + adddays;
                                    // adddays = 4; firstadd += adddays;

                                    // }
                                    // else if (adddays == 4)
                                    // {
                                    // if (Col >= 39) Col = 39;
                                    // workSheet.Cells[Row + 1, Col, Row + 1, (Col + adddays - 1)].Merge = true;
                                    // newdt = dtoprsupval.AsEnumerable().Skip(firstadd).Take(adddays).CopyToDataTable();
                                    // if ((newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Count() > 0))
                                    // {
                                    // workSheet.Cells[Row + 1, Col].Value = newdt.AsEnumerable().Where(x => x.Field<string>("SupervisorName") != null).Select(x => x.Field<string>("SupervisorName")).Distinct().First().ToString() + " (" + newdt.AsEnumerable().Where(x => x.Field<DateTime?>("SupervisorTS") != null).Select(x => x.Field<DateTime?>("SupervisorTS")).Distinct().First().ToString() + " )";
                                    // }
                                    // Col = Col + adddays;
                                    // adddays = 3; firstadd += adddays;
                                    // }
                                    //}
                                    //if (!(Col >= 39))
                                    // workSheet.Cells[Row + 1, Col, Row + 1, 39].Merge = true;
                                    #endregion
                                    string Name = dtoprsupval.Rows[0]["SupervisorName"].ToString();
                                    string Timestamp = dtoprsupval.Rows[0]["SupervisorTS"].ToString();
                                    //workSheet.Cells[Row + 1, 9].Value = dtoprsupval.Rows[0]["SupervisorTS"].ToString() + " ( " +dtoprsupval.Rows[0]["SupervisorTS"].ToString() + " )";
                                    int Col = 9; int tillmerge = 9;
                                    foreach (DataRow dataRow in dtoprsupval.Rows)
                                    {
                                        workSheet.Cells[Row, Col].Value = string.IsNullOrEmpty(dataRow["OperatorName"].ToString()) ? "" : dataRow["OperatorName"].ToString();
                                        if (!(Timestamp.Equals(dataRow["SupervisorTS"].ToString()) && Name.Equals(dataRow["SupervisorName"].ToString())))
                                        {

                                            if (!(string.IsNullOrEmpty(Name) && string.IsNullOrEmpty(Timestamp)))
                                                workSheet.Cells[Row + 1, tillmerge].Value = Name + " ( " + Convert.ToDateTime(Timestamp).ToString("dd-MM-yyyy")+ " )";
                                            if (Col != tillmerge)
                                                workSheet.Cells[Row + 1, tillmerge, Row + 1, Col - 1].Merge = true;

                                            tillmerge = Col;

                                        }
                                        Col++;
                                        if (!string.IsNullOrEmpty(Timestamp))
                                            workSheet.Cells[Row + 1, tillmerge].Value = dataRow["SupervisorName"].ToString() + Name + " ( " + Convert.ToDateTime(Timestamp).ToString("dd-MM-yyyy") + " )";
                                        Timestamp = dataRow["SupervisorTS"].ToString();
                                        Name = dataRow["SupervisorName"].ToString();
                                        ProdHead = string.IsNullOrEmpty(dataRow["ProdHeadName"].ToString()) ? ProdHead : dataRow["ProdHeadName"].ToString();
                                        ProdHeadTS = string.IsNullOrEmpty(dataRow["ProdHeadTS"].ToString()) ? ProdHead : dataRow["ProdHeadTS"].ToString();
                                    }
                                    if (tillmerge != 39)
                                        workSheet.Cells[Row + 1, tillmerge, Row + 1, 39].Merge = true;
                                    Col--;
                                    workSheet.Cells[Row + 2, 9, Row + 2, 39].Merge = true;
                                    workSheet.Row(Row + 1).Height = 36;
                                    workSheet.Row(Row + 2).Height = 36;
                                    if(!string.IsNullOrEmpty(ProdHeadTS))
                                        workSheet.Cells[Row + 2, 9].Value = ProdHead + " ( " + Convert.ToDateTime(ProdHeadTS).ToString("dd-MM-yyyy")+ " )";

                                    workSheet.Cells[Row + 1, 9, Row + 2, 39].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    workSheet.Cells[Row + 1, 9, Row + 2, 39].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    workSheet.Cells[Row, 8, Row + 2, 8].Style.Font.Bold = true;
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Top.Color.SetColor(Color.Black);
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Bottom.Color.SetColor(Color.Black);
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Right.Color.SetColor(Color.Black);
                                    workSheet.Cells[Row, 8, Row + 2, 39].Style.Border.Left.Color.SetColor(Color.Black);
                                    workSheet.Cells[Row, 8, Row, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[Row, 8, Row, 39].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#70AD47"));
                                    workSheet.Cells[Row + 1, 8, Row + 1, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[Row + 1, 8, Row + 1, 39].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFD966"));
                                    workSheet.Cells[Row + 2, 8, Row + 2, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    workSheet.Cells[Row + 2, 8, Row + 2, 39].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4B084"));
                                    workSheet.Cells["C5"].Value = Machine;
                                    workSheet.Cells["G5"].Value = cellID.Equals(string.Empty) ? "ALL" : cellID;
                                    workSheet.Cells["J5"].Value = startTime.ToString("MMM-yyyy");
                                    //workSheet.Name = dtoprsupval.Rows[0]["ShiftName"].ToString();
                                }
                            }

                        }
                    }
                }
                if (Excel.Workbook.Worksheets.Count > 1)
                {
                    Excel.Workbook.Worksheets.Delete(1);
                }

                Excel.SaveAs(newFile);
                Logger.WriteDebugLog("JHAuditReport Report Generated Successfully.");

                if (isDataAvailable)
                {
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                    Logger.WriteDebugLog("JH Audit Report Report Report Exported successfully");
                }
                else
                {
                    Logger.WriteDebugLog("JH Audit Report Report Report not mailed: no data");
                }


            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        public static bool ExportDailyProdDownDaywiseExcelReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
         string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string CompanyName, bool MachineAE)
        {

            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet Wrksht_down = null;
            Excel.Worksheet wrksht = null;
            int pid = 0;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                string src, dst = string.Empty;//Globally Used  

                //----------------------------------------------------Function Local Used Variable-------------------------------------
                string spstr, plantname, exprtpath, rangesel, Catagory, Str, dst_month, mcStr = string.Empty;
                int Mcstart, McEnd, Mcsno, indx_col, Flag_Showexl, Flag_dltshts, r, c, Intrv, Intrv_month, Intrv_indx, actvshtno = 0;
                int x, Out_X, shtno_prd, sht_no, shno = 0;
                double val1, strsht = 0;
                char a = 'F';
                Intrv = 0;

                //---------------------------------------------------------------------------------------------------------------------
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\SM_Daily_ProdDown_daywise.xls";
                DateTime DT = DateTime.Now.AddDays(-1);
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }
                plantname = (plantid == "") ? "All Plant" : plantid;
                indx_col = AccessReportData.MaxHourIdShift();

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"Production_Down_monthwise_" + plantname + "_" + string.Format("{0:dd_MMM_yyyy_HH_mm}", DateTime.Now.AddDays(-1)) + ".xls";//string.Format("{0:hh-mm-ss MMM-yyyy}", DT) + ".xls";
                if (!File.Exists(dst))
                {
                    File.Copy(src, dst, true);
                }
                Thread.Sleep(1000);
                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;

                int b = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                if (!File.Exists(dst))
                {
                    return false;
                }

                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                for (Out_X = 1; Out_X < 31; Out_X++)
                {
                    for (x = 0; x < 1; x++)
                    {
                        if (wrkbk.Sheets.Count <= 64)
                        {
                            Excel.Worksheet exWS = (Excel.Worksheet)wrkbk.Worksheets.get_Item(wrkbk.Sheets.Count);
                            Excel.Worksheet exWS1 = (Excel.Worksheet)wrkbk.Worksheets.get_Item(wrkbk.Sheets.Count - 1);
                            exWS1.Name = Out_X.ToString() + "P";
                            exWS1.Copy(misValue, exWS);

                            exWS = (Excel.Worksheet)wrkbk.Worksheets.get_Item(wrkbk.Sheets.Count);
                            exWS1 = (Excel.Worksheet)wrkbk.Worksheets.get_Item(wrkbk.Sheets.Count - 1);
                            exWS1.Name = Out_X.ToString() + "D";
                            exWS1.Copy(misValue, exWS);
                        }
                    }
                }

                shno = 1;
                string str = string.Format("{0:dd}", DT);
                val1 = (int.Parse(str)) * 2 + (Intrv * 2);


                if (wrkbk.Sheets.Count >= val1)
                {
                    strsht = ((int.Parse(str)) * 2) - 2;
                }
                else
                {
                    Logger.WriteDebugLog("Tempalte got corrupted " + src);
                    xlApp.Quit();
                    releaseObject(xlApp);
                    releaseObject(wrkbk);
                    return false;
                }
                r = 3;
                c = 1;

                SqlDataReader rs = AccessReportData.ProdDownReport(DateTime.Now.AddDays(-1), DateTime.Now, plantid, MachineId, "Production");

                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(strsht + 1);
                Excel.Range workSheet_range = null;

                workSheet_range = wrksht.get_Range("A3", "w4000");
                workSheet_range.Interior.ColorIndex = 0;
                workSheet_range.ClearContents();

                wrksht.Cells[1, 1] = "Production Report On  " + string.Format("{0:dd-MMM-yyyy}", DT) + " ( " + DT.DayOfWeek.ToString() + " )";
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(strsht + 1);
                wrksht.Activate();

                if (rs.Read())
                {
                    workSheet_range = wrksht.get_Range("F1", "Q1");
                    workSheet_range.EntireColumn.Hidden = false;
                    if (indx_col > 0 && indx_col < 12)
                    {
                        rangesel = Convert.ToChar((int)a + indx_col) + "3";
                        workSheet_range = wrksht.get_Range(rangesel, "Q65536");
                        workSheet_range.EntireColumn.Hidden = true;
                    }
                    rangesel = "";

                    Mcstart = r;
                    Mcsno = 1;

                    Flag_dltshts = 0;
                    Flag_Showexl = 1;

                    wrksht.Cells[r, c] = rs["ShiftName"];
                    c = c + 1;
                    wrksht.Cells[r, c] = rs["MachineID"];
                    mcStr = rs["MachineID"].ToString();
                    c = c + 1;
                    wrksht.Cells[r, c] = rs["ComponentId"];
                    c = c + 1;
                    wrksht.Cells[r, c] = "";
                    c = c + 1;
                    wrksht.Cells[r, c] = rs["Operationno"];
                    c = c + 1;
                    wrksht.Cells[r, c] = rs["Operatorid"];
                    c = c + 1;
                    wrksht.Cells[r, c] = rs["Actual"];
                    c = c + 1;
                    wrksht.Cells[r, 18] = rs["Downtime"];
                    wrksht.Cells[r, 19] = rs["Hourlytarget"];
                    wrksht.Cells[r, 20] = rs["Target"];
                    wrksht.Cells[r, 21] = rs["ShftactualCount"];

                    while (rs.Read())
                    {
                        if (rs.HasRows)
                        {

                            if (wrksht.Cells[r, 2] != rs["MachineID"])
                            {
                                if (Mcsno % 2 == 1)
                                {
                                    rangesel = "A" + Mcstart + ":W" + r;
                                    workSheet_range = wrksht.get_Range("A" + Mcstart, "W" + r);
                                    workSheet_range.Interior.ColorIndex = 40;
                                    workSheet_range = null;
                                }
                                Mcstart = r + 1;
                                Mcsno = Mcsno + 1;
                            }

                            if ((wrksht.Cells[r, 1] == rs["ShiftName"]) && (wrksht.Cells[r, 2] == rs["MachineID"]) && (wrksht.Cells[r, 3] == rs["ComponentID"]) && (wrksht.Cells[r, 4] == rs["OperationNo"]) && (wrksht.Cells[r, 5] == rs["Operatorid"]))
                            {
                                wrksht.Cells[r, c] = rs["Actual"];
                                c = c + 1;
                            }
                            else
                            {
                                r = r + 1;
                                c = 1;
                                wrksht.Cells[r, c] = rs["ShiftName"];
                                c = c + 1;
                                wrksht.Cells[r, c] = rs["MachineID"];
                                mcStr = rs["MachineID"].ToString();
                                c = c + 1;
                                wrksht.Cells[r, c] = rs["ComponentID"];
                                c = c + 1;
                                wrksht.Cells[r, c] = rs["OperationNo"];
                                c = c + 1;
                                wrksht.Cells[r, c] = rs["Operatorid"];
                                c = c + 1;
                                wrksht.Cells[r, c] = rs["Actual"];
                                c = c + 1;
                                wrksht.Cells[r, 18] = rs["Downtime"];
                                wrksht.Cells[r, 19] = rs["Hourlytarget"];
                                wrksht.Cells[r, 20] = rs["Target"];
                                wrksht.Cells[r, 21] = rs["ShftactualCount"];
                            }
                        }
                    }
                    wrksht.Columns.AutoFit();
                }
                else
                {
                    Logger.WriteDebugLog("Record not found.");
                    wrksht = null;
                }

                Mcsno = 0;
                Mcstart = 0;
                if (r > 3 && (Mcsno % 2 == 1))
                {
                    rangesel = "A" + Mcstart + ":W" + r;
                    workSheet_range = wrksht.get_Range("A" + Mcstart, "W" + r);
                    workSheet_range.Interior.ColorIndex = 40;
                    workSheet_range = null;
                }
                if (rs != null)
                {
                    rs.Close();
                }

                r = 3;
                c = 1;

                mcStr = "";
                rs = AccessReportData.ProdDownReport(DateTime.Now.AddDays(-1), DateTime.Now, plantid, MachineId, "Down");
                sht_no = (int)strsht + 2;

                Wrksht_down = (Excel.Worksheet)wrkbk.Worksheets.get_Item(sht_no);

                workSheet_range = Wrksht_down.get_Range("A3", "w65000");
                workSheet_range.ClearContents();
                workSheet_range.Interior.ColorIndex = 0;
                workSheet_range = null;

                Wrksht_down.Cells[1, 1] = "Down Report On  " + string.Format("{0:dd-MMM-yyyy}", DateTime.Now.AddDays(-1).ToString()) + " ( " + DateTime.Now.AddDays(-1).DayOfWeek.ToString() + " )";

                if (rs.Read())
                {
                    Mcstart = r;
                    Mcsno = 1;

                    Wrksht_down.Cells[r, c] = rs["MachineID"];
                    mcStr = rs["MachineID"].ToString();
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["ComponentID"];
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["Operationno"];
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["OperatorID"];
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["Starttime"];
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["Endtime"];
                    c = c + 1;

                    workSheet_range = (Excel.Range)Wrksht_down.Rows[r, misValue];
                    workSheet_range.Font.Bold = false;
                    workSheet_range = null;

                    Wrksht_down.Cells[r, c] = rs["Downtime"];
                    Wrksht_down.Cells[r + 1, c] = rs["Mcwisecnt"];

                    workSheet_range = (Excel.Range)Wrksht_down.Rows[r + 1, misValue];
                    workSheet_range.Font.Bold = true;
                    workSheet_range = null;

                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["DownID"];
                    c = c + 1;
                    Wrksht_down.Cells[r, c] = rs["Remarks"];

                    while (rs.Read())
                    {
                        if (rs.HasRows)
                        {
                            r = r + 1;
                            c = 1;
                            if (mcStr != rs["MachineID"].ToString())
                            {
                                if (Mcsno % 2 == 1)
                                {
                                    rangesel = "A" + Mcstart + ":I" + r;
                                    workSheet_range = Wrksht_down.get_Range("A" + Mcstart, "I" + r);
                                    workSheet_range.Interior.ColorIndex = 40;
                                    workSheet_range = null;
                                }
                                r = r + 1;
                                Mcstart = r;
                                Mcsno = Mcsno + 1;
                            }

                            Wrksht_down.Cells[r, c] = rs["MachineID"];
                            mcStr = rs["MachineID"].ToString();
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["ComponentID"];
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["Operationno"];
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["OperatorID"];
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["Starttime"];
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["Endtime"];
                            c = c + 1;

                            workSheet_range = (Excel.Range)Wrksht_down.Rows[r, misValue];
                            workSheet_range.Font.Bold = false;
                            workSheet_range = null;

                            Wrksht_down.Cells[r, c] = rs["Downtime"];
                            Wrksht_down.Cells[r + 1, c] = rs["Mcwisecnt"];

                            workSheet_range = (Excel.Range)Wrksht_down.Rows[r + 1, misValue];
                            workSheet_range.Font.Bold = true;
                            workSheet_range = null;

                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["DownID"];
                            c = c + 1;
                            Wrksht_down.Cells[r, c] = rs["Remarks"];
                        }

                    }
                    if (rs != null)
                    {
                        rs.Close();
                    }
                    Wrksht_down.Columns.AutoFit();
                }
                else
                {
                    Logger.WriteDebugLog("Record not found.");
                }

                if (r > 3 && Mcsno % 2 == 1)
                {
                    rangesel = "A" + Mcstart + ":I" + r + 1;
                    workSheet_range = Wrksht_down.get_Range("A" + Mcstart, "I" + r + 1);
                    workSheet_range.Interior.ColorIndex = 0;
                    workSheet_range = null;
                }
                wrkbk.Close(true, misValue, misValue);
                if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");

                dst = GetZipFile(dst);
                Thread.Sleep(7000);

                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                if (xlApp != null)
                {
                    releaseObject(wrkbk);
                    releaseObject(Wrksht_down);
                    releaseObject(wrksht);
                    releaseObject(xlApp);
                }
                if (pid != 0) KillSpecificExcelFileProcess(pid);
            }
            return true;
        }

        internal static void ExportPMPhantomCellReport(string strReportFile, string ExportPath, string ExportFileName, string Machine, DateTime startTime, DateTime endTime, string plnt, string Cell, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string RunReportForEvery)
        {
            string dst = string.Empty;
            bool isDataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("PMReport Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath, string.Format("PM_Phantom_Report_{0:ddMMMyyyy_HHmmss}.xlsx", DateTime.Now));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage Excel = new ExcelPackage(newFile, true);
                var workBook = Excel.Workbook;
                var workSheet = workBook.Worksheets[1];
                System.Data.DataTable dtList = AccessReportData.GetPMReportData(Cell, startTime, Machine, plnt, endTime);
                
                if (dtList != null && dtList.Rows.Count > 0)
                {
                    isDataAvailable = true;

                    //var machineList = dtList.AsEnumerable().Select(x => x.Field<string>("Machine")).Distinct().ToList();
                    //int sheet = 1;
                    //foreach (var machine in machineList)
                    //{
                    //    if (sheet == 1)
                    //    {
                    //        workSheet.Name = machine;

                    //    }
                    //    else
                    //    {
                    //        Excel.Workbook.Worksheets.Add(machine, workSheet);
                    //    }
                    //    sheet++;
                    //}

                    //foreach(string machine in machineList)
                    //{
                        
                    //    System.Data.DataTable dtDistMachinedata = dtList.AsEnumerable().Where(x => x.Field<string>("MachineID").Equals(machine)).CopyToDataTable();
                    //}

                    int row = 6; int slno = 1;                    
                    int Firstcol = 4, lastcol = 5;
                    
                    List<System.Data.DataTable> Category = dtList.AsEnumerable().GroupBy(z => z.Field<string>("Category")).Select(x => x.CopyToDataTable()).Distinct().ToList();
                    int col = 4;
                    string month = Convert.ToDateTime(dtList.Columns[3].ColumnName).ToString("MMMM");
                    for (int i = 3; i < dtList.Columns.Count; i++)
                    {
                        if (!(month == Convert.ToDateTime(dtList.Columns[i].ColumnName).ToString("MMMM")))
                        {
                            if (Firstcol != col)
                            {
                                workSheet.Cells[row - 1, Firstcol, row - 1, col - 1].Merge = true;
                                workSheet.Cells[row - 1, Firstcol].Value = month;
                            }
                            Firstcol = col;
                        }
                        workSheet.Cells[row, col++].Value = (dtList.Columns[i].ColumnName);
                    }
                    workSheet.Cells[row - 1, Firstcol].Value = Convert.ToDateTime(dtList.Columns[dtList.Columns.Count - 1].ColumnName).ToString("MMMM");
                    foreach (System.Data.DataTable dtcat in Category)
                    {
                        if (dtcat != null && dtcat.Rows.Count > 0)
                        {
                            workSheet.Cells[row, 1].Value = dtcat.Rows[0]["Category"].ToString();
                            workSheet.Cells[row, 1, row, 2].Merge = true;
                            workSheet.Cells[row++, 1].Style.Font.Bold = true;
                            row++;

                            col = 1;
                            for (int i = 0; i < dtcat.Rows.Count; i++)
                            {
                                col = 1;
                                workSheet.Cells[row, col++].Value = slno++;
                                workSheet.Cells[row, col++].Value = dtcat.Rows[i]["Items"];
                                workSheet.Cells[row, col++].Value = "Cycles";
                               // workSheet.Cells[row, col++].Value = dtcat.Rows[i]["Remarks"];
                                for (int j = 3; j < dtcat.Columns.Count; j++)
                                {
                                    workSheet.Cells[row, col].Value = dtcat.Rows[i][j];
                                    if ((!dtcat.Rows[i][j].ToString().Contains("NotOK")))
                                    {
                                        workSheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#90EE90"));
                                    }
                                    else
                                    {
                                        workSheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        workSheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#E9323C"));
                                    }
                                    col++;
                                }
                                row++;
                            }
                            lastcol = col;
                        }
                    }
                    if (lastcol < 21)
                        lastcol = 22;
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Top.Color.SetColor(Color.Black);
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Left.Color.SetColor(Color.Black);
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Right.Color.SetColor(Color.Black);
                    workSheet.Cells[5, 1, row, lastcol - 1].Style.Border.Bottom.Color.SetColor(Color.Black);
                    workSheet.Cells["T3"].Value = startTime.ToString("dd-MM-yyyy");
                    //workSheet.Cells["T4"].Value = endTime.ToString("dd-MM-yyyy");
                    workSheet.Cells["F3"].Value = string.IsNullOrEmpty(plnt) ? "All" : plnt;
                    workSheet.Cells["D3"].Value = string.IsNullOrEmpty(Machine) ? "All" : Machine;
                    workSheet.Cells["E4"].Value = "Cell ID:";
                    workSheet.Cells["F4"].Value = "Phantom";
                }
                Excel.SaveAs(newFile);
                Logger.WriteDebugLog("PM Report has been generated successfully.");

                if (isDataAvailable)
                {
                    Logger.WriteDebugLog("PM Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportFileName);
                }
                else
                {
                    Logger.WriteDebugLog("PM Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
                throw;
            }
        }

        public static bool ExportDNCUsageReport(string strReportFile, string ExportPath, string ExportedReportFile,
           int ExportType, int DayBefores, string reportMode, string MachineId, string operators, string sttime,
           string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string CompanyName, bool MachineAE)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet wrksht = null;
            int pid = 0;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                string src = string.Empty, dst = string.Empty;//Globally Used  

                SqlDataReader rs = AccessReportData.DNCUsageReport(DateTime.Parse(sttime), DateTime.Parse(ndtime), MachineId);

                switch (ExportType)
                {
                    case 0://.xls
                        string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        if (reportMode.ToUpper() == "DAY")
                            src = APath + @"\Reports\DNCUsageReport_TemplateDayWise.xls";
                        if (reportMode.ToUpper() == "MONTH")
                            src = APath + @"\Reports\DNCUsageReport_TemplateMonthWise.xls";

                        if (!File.Exists(src))
                        {
                            Logger.WriteDebugLog("Template is not found on " + src);
                            return false;
                        }

                        if (!Directory.Exists(ExportPath))
                        {
                            Directory.CreateDirectory(ExportPath);
                        }
                        dst = ExportPath + @"CamShaft_DNCUsageReport_" + string.Format("{0:dd_MMM_yyyy}_{1}wise", Convert.ToDateTime(sttime), reportMode) + ".xls";
                        if (!File.Exists(dst))
                        {
                            File.Copy(src, dst, true);
                        }

                        Thread.Sleep(1000);
                        if (!File.Exists(dst))
                        {
                            return false;
                        }

                        try
                        {
                            xlApp = new Excel.ApplicationClass();
                            xlApp.DisplayAlerts = false;
                            int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                            wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                            wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item("DNC Usage Report");
                            int irow = 5, rowMergeStart = 5;
                            int icol = 1;
                            DateTime dt = DateTime.MinValue;
                            int i = 1;
                            while (rs.Read())
                            {
                                if (reportMode.ToUpper() == "MONTH")
                                {
                                    if (dt == DateTime.MinValue || dt != Convert.ToDateTime(rs["Startdate"]))
                                    {
                                        if (dt != DateTime.MinValue)
                                        {
                                            // wrksht.get_Range((Excel.Range)wrksht.Cells[rowMergeStart, 7], (Excel.Range)wrksht.Cells[irow - 1, 7]).Merge(false);
                                            wrksht.get_Range((Excel.Range)wrksht.Cells[irow, 1], (Excel.Range)wrksht.Cells[irow, 10]).Interior.ColorIndex = 45;
                                            irow++;
                                        }
                                        dt = Convert.ToDateTime(rs["Startdate"]);
                                        rowMergeStart = irow;
                                    }
                                }
                                wrksht.Cells[irow, icol++] = i;
                                wrksht.Cells[irow, icol++] = rs["UserName"];
                                wrksht.Cells[irow, icol++] = rs["ClientName"];
                                wrksht.Cells[irow, icol++] = rs["MachineID"];
                                wrksht.Cells[irow, icol++] = rs["LogMessage"];
                                wrksht.Cells[irow, icol++] = rs["ProgramID"];
                                if (reportMode.ToUpper() == "MONTH")
                                {
                                    wrksht.Cells[irow, icol++] = rs["TransferStart"];
                                }
                                wrksht.Cells[irow, icol++] = rs["TransferStart"];
                                if (reportMode.ToUpper() == "MONTH")
                                {
                                    wrksht.Cells[irow, 9] = rs["TransferEnd"];
                                    wrksht.Cells[irow, 10] = rs["QTY"];
                                }
                                else
                                {
                                    wrksht.Cells[irow, 8] = rs["TransferEnd"];
                                    wrksht.Cells[irow, 9] = rs["QTY"];
                                }

                                i = i + 1;
                                irow++;
                                icol = 1;
                            }
                            if (rs != null)
                            {
                                rs.Close();
                            }
                            if (reportMode.ToUpper() == "MONTH")
                            {
                                //if (irow > 5) wrksht.get_Range((Excel.Range)wrksht.Cells[rowMergeStart, 7], (Excel.Range)wrksht.Cells[irow - 1, 7]).Merge(false);
                                //else wrksht.get_Range((Excel.Range)wrksht.Cells[rowMergeStart, 7], (Excel.Range)wrksht.Cells[irow, 7]).Merge(false);
                                if (irow <= 5) wrksht.get_Range((Excel.Range)wrksht.Cells[rowMergeStart, 7], (Excel.Range)wrksht.Cells[irow, 7]).Merge(false);
                            }
                            wrksht.Cells[2, 3] = string.Format("{0:dd-MMM-yyyy hh:mm:ss tt}", sttime);
                            wrksht.Cells[2, 5] = string.Format("{0:dd-MMM-yyyy hh:mm:ss tt}", ndtime);
                            wrksht.Columns.AutoFit();
                            wrkbk.Close(true, misValue, misValue);
                            if (xlApp != null) xlApp.Quit();
                            Logger.WriteDebugLog("Report generated sucessfully.");
                            if (irow > 5)
                            {
                                Thread.Sleep(1000);
                                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                            }
                            else
                            {
                                Logger.WriteDebugLog("Data not found for selected time period. No need to send email.");
                            }
                        }
                        catch (Exception ex)
                        {

                            Logger.WriteErrorLog(ex.ToString());
                        }
                        finally
                        {
                            if (xlApp != null)
                            {
                                releaseObject(xlApp);
                                releaseObject(wrkbk);
                                releaseObject(wrksht);
                            }
                            if (pid != 0) KillSpecificExcelFileProcess(pid);
                        }
                        break;
                    case 1://Html
                        string DocWrite;
                        dst = ExportPath + @"CamShaft_DNCUsageReport_" + string.Format("{0:yyyyMMdd_HHmmss}", ndtime) + ".Html";

                        DocWrite = "<head><title>From Date</title> </head> <body> <table border=1 cellspacing=1  style=border-collapse: collapse bordercolor=#111111 width=100% id=AutoNumber1>";
                        DocWrite = DocWrite + "<tr>    <td width=114% colspan=8 bgcolor=#FF9933>   <p align=center><b>DNC USAGE REPORT</b></td> ";
                        DocWrite = DocWrite + "</tr>  <tr>    <td width=19% colspan=2><b>From Date:</b></td> <td nowrap width=13%>" + string.Format("{0:DD-MMM-YYYY hh:mm:ss tt}", sttime) + "</td><td width=12% colspan=1><b>To Date:</b></td><td nowrap width=23%>" + string.Format("{0:DD-MMM-YYYY hh:mm:ss tt}", ndtime) + "</td>";
                        DocWrite = DocWrite + "</td> <td width=1%>&nbsp;</td>    <td width=12%>&nbsp;</td> <td width=12%>&nbsp;</td> <tr>";

                        WriteInToFile(dst, DocWrite);

                        DocWrite = string.Empty;

                        WriteInToFile(dst, "<tr bgcolor=#FF9933>");
                        WriteInToFile(dst, "<td width=4%><b><font size=2>S.NO.</font></b></td>");
                        WriteInToFile(dst, "<td width=1%><b>User Name</b></td>");
                        WriteInToFile(dst, "<td width=13%><b>Client Name</b></td>");
                        WriteInToFile(dst, "<td width= 15%><b>Machine ID</b></td>");
                        WriteInToFile(dst, "<td width=9%><b>Log Message</b></td>");
                        WriteInToFile(dst, "<td width=11%><b>Program ID</b></td>");
                        WriteInToFile(dst, "<td width=60%><b>Time Stamp-Transfer Start</b></td>");
                        WriteInToFile(dst, "<td width=60%><b>Time Stamp-Transfer End</b></td>");
                        WriteInToFile(dst, "<td width=1%><b>QTY</b></td>");
                        WriteInToFile(dst, "</tr>");

                        int Row = 1;
                        while (rs.Read())
                        {
                            WriteInToFile(dst, "<tr >");
                            WriteInToFile(dst, "<td width=4%>" + Row + "</td>");
                            WriteInToFile(dst, "<td width=1%>" + rs["UserName"] + "</td>");
                            WriteInToFile(dst, "<td width=13%>" + rs["ClientName"] + "</td>");
                            WriteInToFile(dst, "<td width= 15%>" + rs["MachineID"] + "</td>");
                            WriteInToFile(dst, "<td width=9%>" + rs["LogMessage"] + "</td>");
                            WriteInToFile(dst, "<td width=11%>" + rs["ProgramID"] + "</td>");
                            WriteInToFile(dst, "<td width=60%>" + rs["TransferStart"] + "</td>");
                            WriteInToFile(dst, "<td width=60%>" + rs["TransferEnd"] + "</td>");
                            WriteInToFile(dst, "<td width=1%>" + rs["QTY"] + "</td>");
                            WriteInToFile(dst, "</tr>");
                            Row = Row + 1;
                        }

                        WriteInToFile(dst, "</table></body></html>");

                        Logger.WriteDebugLog("Report generated sucessfully");

                        Thread.Sleep(2000);
                        SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                        break;
                    case 2://.Csv
                        dst = ExportPath + @"CamShaft_DNCUsageReport_" + string.Format("{0:yyyyMMdd_HHmmss}", ndtime) + ".csv";

                        WriteInToFile(dst, "DNC USAGE REPORT");
                        WriteInToFile(dst, "FromDate:," + string.Format("{0:DD-MMM-YYYY hh:mm:ss tt}", sttime) + "," + "ToDate:," + string.Format("{0:DD-MMM-YYYY hh:mm:ss tt}", ndtime));
                        WriteInToFile(dst, "Sl.No,UserName,ClientName,MachineID,Logmessage,ProgramID,TransferStart,TransferEnd,QTY");

                        Row = 1;
                        while (rs.Read())
                        {
                            WriteInToFile(dst, Row + "," + rs["UserName"] + "," + rs["ClientName"] + "," + rs["MachineID"] + "," + rs["LogMessage"] + "," + rs["ProgramID"] + "," + rs["TransferStart"] + "," + rs["TransferEnd"] + "," + rs["QTY"]);
                            Row = Row + 1;
                        }
                        if (rs != null)
                        {
                            rs.Close();
                        }

                        Logger.WriteDebugLog("Report generated sucessfully.");
                        Thread.Sleep(2000);
                        SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(" Report generation Failed Error: " + ex.ToString());
                return false;
            }

            finally
            {
                if (xlApp != null)
                {
                    releaseObject(xlApp);
                    releaseObject(wrkbk);
                    releaseObject(wrksht);
                    if (pid != 0) KillSpecificExcelFileProcess(pid);
                }

            }
            return true;
        }

        internal static void ExportSAPOEEReportAdvik(DateTime fromDate, DateTime toDate, string ReportFileName, string ExportPath, string ExportFileName, string Shift, string Machine, string plnt, string cell, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string RunReportForEvery)
        {
            bool isDataAvailable = false;
            //Logger.WriteDebugLog("SAP OEE Report");
            try
            {
                string Source = string.Empty, dst = string.Empty, Template = string.Empty, time = string.Empty;
                int row = 8;
                string Filename = "AdvikSAPOEEReport.xlsx";
                if (!File.Exists(ReportFileName))
                {
                    Logger.WriteDebugLog("AdvikSAPOEEReport Template is not found on " + ReportFileName);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SAP_OEE_{1}_Report_{0:ddMMMyyyy_HHmmss}.xlsx", fromDate, RunReportForEvery));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }

                File.Copy(ReportFileName, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage Excel = new ExcelPackage(newFile, true);
                var ws = Excel.Workbook.Worksheets[1];

                ws.Cells["B4"].Value = fromDate.ToString("dd-MM-yyyy");
                ws.Cells["E4"].Value = toDate.ToString("dd-MM-yyyy");
                ws.Cells["H4"].Value = Machine;
                ws.Cells["J4"].Value = Shift;

                Logger.WriteDebugLog("Time Interval - From : " + fromDate.ToString("dd-MM-yyyy HH:mm:ss") + " To : " + toDate.ToString("dd-MM-yyyy HH:mm:ss"));
                System.Data.DataTable HeaderDowm = new System.Data.DataTable(); int col = 17;
                System.Data.DataTable SAPOEEDATATABLE = AccessReportData.GetSAPOEEData(fromDate, toDate, Machine, Shift, plnt, cell, out HeaderDowm);

                if (HeaderDowm != null && HeaderDowm.Rows.Count > 0 && SAPOEEDATATABLE != null && SAPOEEDATATABLE.Rows.Count > 0)
                {
                    //FileInfo newFile = new FileInfo(Source);
                    //ExcelPackage Excel = new ExcelPackage(newFile, true);
                    isDataAvailable = true;
                    
                    foreach (DataRow rows in HeaderDowm.Rows)
                    {
                        ws.Cells[7, col].Value = rows["DownDescription"]; col++;
                    }
                    foreach (DataRow rows in SAPOEEDATATABLE.Rows)
                    {
                        ws.Cells[row, 1].Value = rows["Pdate"];
                        ws.Cells[row, 2].Value = rows["WorkCenter"];
                        ws.Cells[row, 3].Value = rows["OrderType"];
                        ws.Cells[row, 4].Value = rows["EquipmentName"];
                        ws.Cells[row, 5].Value = rows["EquipmentNo"];
                        ws.Cells[row, 6].Value = rows["Materialcode"];
                        ws.Cells[row, 7].Value = rows["Shift"];
                        ws.Cells[row, 8].Value = string.IsNullOrEmpty(rows["TotalAvailableTime"].ToString()) ? 0 : Convert.ToDouble(rows["TotalAvailableTime"].ToString());
                        ws.Cells[row, 9].Value = string.IsNullOrEmpty(rows["PlannedAvailableTime"].ToString()) ? 0 : Convert.ToDouble(rows["PlannedAvailableTime"].ToString());
                        ws.Cells[row, 10].Value = string.IsNullOrEmpty(rows["TotalPlannedQty"].ToString()) ? 0 : Convert.ToDouble(rows["TotalPlannedQty"].ToString());
                        ws.Cells[row, 11].Value = string.IsNullOrEmpty(rows["TotalActualQty"].ToString()) ? 0 : Convert.ToDouble(rows["TotalActualQty"].ToString());
                        ws.Cells[row, 12].Value = string.IsNullOrEmpty(rows["TotalOKQty"].ToString()) ? 0 : Convert.ToDouble(rows["TotalOKQty"].ToString());
                        ws.Cells[row, 13].Value = string.IsNullOrEmpty(rows["Reworkqty"].ToString()) ? 0 : Convert.ToDouble(rows["Reworkqty"].ToString());
                        ws.Cells[row, 14].Value = string.IsNullOrEmpty(rows["Scrapqty"].ToString()) ? 0 : Convert.ToDouble(rows["Scrapqty"].ToString());
                        ws.Cells[row, 15].Value = string.IsNullOrEmpty(rows["TotalDefectiveQty"].ToString()) ? 0 : Convert.ToDouble(rows["TotalDefectiveQty"].ToString());
                        ws.Cells[row, 16].Value = string.IsNullOrEmpty(rows["StdCycleTime"].ToString()) ? 0 : Convert.ToDouble(rows["StdCycleTime"].ToString());

                        ws.Cells[row, 17].Value =string.IsNullOrEmpty(rows["A"].ToString())? 0 : Convert.ToDouble(rows["A"].ToString());
                        ws.Cells[row, 18].Value =string.IsNullOrEmpty(rows["B"].ToString())? 0 : Convert.ToDouble(rows["B"].ToString());
                        ws.Cells[row, 19].Value =string.IsNullOrEmpty(rows["C"].ToString())? 0 : Convert.ToDouble(rows["C"].ToString());
                        ws.Cells[row, 20].Value =string.IsNullOrEmpty(rows["D"].ToString())? 0 : Convert.ToDouble(rows["D"].ToString());
                        ws.Cells[row, 21].Value =string.IsNullOrEmpty(rows["E"].ToString())? 0 : Convert.ToDouble(rows["E"].ToString());
                        ws.Cells[row, 22].Value =string.IsNullOrEmpty(rows["F"].ToString())? 0 : Convert.ToDouble(rows["F"].ToString());
                        ws.Cells[row, 23].Value =string.IsNullOrEmpty(rows["G"].ToString())? 0 : Convert.ToDouble(rows["G"].ToString());
                        ws.Cells[row, 24].Value =string.IsNullOrEmpty(rows["H"].ToString())? 0 : Convert.ToDouble(rows["H"].ToString());
                        ws.Cells[row, 25].Value =string.IsNullOrEmpty(rows["I"].ToString())? 0 : Convert.ToDouble(rows["I"].ToString());
                        ws.Cells[row, 26].Value =string.IsNullOrEmpty(rows["J"].ToString())? 0 : Convert.ToDouble(rows["J"].ToString());
                        ws.Cells[row, 27].Value =string.IsNullOrEmpty(rows["K"].ToString())? 0 : Convert.ToDouble(rows["K"].ToString());
                        ws.Cells[row, 28].Value =string.IsNullOrEmpty(rows["L"].ToString())? 0 : Convert.ToDouble(rows["L"].ToString());
                        ws.Cells[row, 29].Value =string.IsNullOrEmpty(rows["M"].ToString())? 0 : Convert.ToDouble(rows["M"].ToString());
                        ws.Cells[row, 30].Value =string.IsNullOrEmpty(rows["N"].ToString())? 0 : Convert.ToDouble(rows["N"].ToString());
                        ws.Cells[row, 31].Value =string.IsNullOrEmpty(rows["O"].ToString())? 0 : Convert.ToDouble(rows["O"].ToString());
                        ws.Cells[row, 32].Value =string.IsNullOrEmpty(rows["P"].ToString())? 0 : Convert.ToDouble(rows["P"].ToString());
                        ws.Cells[row, 33].Value =string.IsNullOrEmpty(rows["Q"].ToString())? 0 : Convert.ToDouble(rows["Q"].ToString());
                        ws.Cells[row, 34].Value =string.IsNullOrEmpty(rows["R"].ToString())? 0 : Convert.ToDouble(rows["R"].ToString());
                        ws.Cells[row, 35].Value =string.IsNullOrEmpty(rows["S"].ToString())? 0 : Convert.ToDouble(rows["S"].ToString());
                        ws.Cells[row, 36].Value = string.IsNullOrEmpty(rows["T"].ToString()) ? 0 : Convert.ToDouble(rows["T"].ToString());
                        ws.Cells[row, 37].Value = string.IsNullOrEmpty(rows["TotalLoss"].ToString()) ? 0 : Convert.ToDouble(rows["TotalLoss"].ToString());
                        ws.Cells[row, 38].Value = string.IsNullOrEmpty(rows["AEffy"].ToString())? 0 : Convert.ToDouble(rows["AEffy"].ToString());
                        ws.Cells[row, 39].Value = string.IsNullOrEmpty(rows["PEffy"].ToString())? 0 : Convert.ToDouble(rows["PEffy"].ToString());
                        ws.Cells[row, 40].Value = string.IsNullOrEmpty(rows["QEffy"].ToString()) ? 0 : Convert.ToDouble(rows["QEffy"].ToString());
                        ws.Cells[row, 41].Value = string.IsNullOrEmpty(rows["OEffy"].ToString()) ? 0 : Convert.ToDouble(rows["OEffy"].ToString());
                        row++;
                    }

                    for (int k = 8; k <= 37; k++)
                    {
                        string Index = GetExcelColumnName(k + 1);
                        string formula = "=SUM(" + Index + 8 + ":" + Index + (row - 1) + ")";
                        ws.Cells[row, (k + 1)].Formula = formula;
                    }

                    for (int k = 37; k <= 40; k++)
                    {
                        string Index = GetExcelColumnName(k + 1);
                        string formula = "=AVERAGE(" + Index + 8 + ":" + Index + (row - 1) + ")";
                        ws.Cells[row, (k + 1)].Formula = formula;
                    }
                    ws.Cells[4, 1, row, 41].AutoFitColumns();
                    ws.Cells[8, 1, row, 41].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[8, 1, row, 41].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[8, 1, row, 41].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[8, 1, row, 41].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[8, 1, row, 41].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[8, 1, row, 41].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[8, 1, row, 41].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[8, 1, row, 41].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Calculate();
                }

                Excel.SaveAs(newFile);
                Logger.WriteDebugLog("SAP-OEE Report Generated Successfully.");

                if (isDataAvailable)
                {                    
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportFileName);
                    Logger.WriteDebugLog("SAP-OEE Report Report Exported successfully");
                }
                else
                {
                    Logger.WriteDebugLog("SAP-OEE Report Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }


        internal static void ExportDowntimeQualificationReportAdvik(DateTime fromDate, DateTime toDate, string ReportFileName, string ExportPath, string ExportFileName, string Machine, string plnt, string cell, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string RunReportForEvery)
        {
            bool isDataAvailable = false;
            //Logger.WriteDebugLog("SAP OEE Report");
            try
            {
                string Source = string.Empty, dst = string.Empty, Template = string.Empty, time = string.Empty;
                int row = 7,rowsCount=1;
                if (!File.Exists(ReportFileName))
                {
                    Logger.WriteDebugLog("DowntimeQualificationReport Template is not found on " + ReportFileName);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("Downtime_Qualification_{1}_Report_{0:ddMMMyyyy_HHmmss}.xlsx", fromDate, RunReportForEvery));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }

                File.Copy(ReportFileName, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage Excel = new ExcelPackage(newFile, true);
                var ws = Excel.Workbook.Worksheets[1];

                ws.Cells["B4"].Value = fromDate.ToString("dd-MM-yyyy HH:mm:ss");
                ws.Cells["D4"].Value = toDate.ToString("dd-MM-yyyy HH:mm:ss");
                ws.Cells["F4"].Value = string.IsNullOrEmpty(plnt) ? "ALL" : plnt;
                ws.Cells["H4"].Value = string.IsNullOrEmpty(cell) ? "ALL" : cell;

                System.Data.DataTable dtDowntimeQlf = AccessReportData.GetDowntimeQualificationData(plnt, Machine, cell, fromDate, toDate);

                if (dtDowntimeQlf != null && dtDowntimeQlf.Rows.Count > 0)
                {
                    isDataAvailable = true;

                    foreach (DataRow rows in dtDowntimeQlf.Rows)
                    {
                        //ws.Cells[row, 1].Value = rowsCount;
                        //ws.Cells[row, 2].Value = rows["id"];
                        ws.Cells[row, 1].Value = rows["machineid"] is DBNull ? string.Empty : rows["machineid"].ToString();
                        ws.Cells[row, 2].Value = rows["downid"] is DBNull ? string.Empty : rows["downid"].ToString();
                        ws.Cells[row, 3].Value = rows["sttime"] is DBNull ? string.Empty : rows["sttime"].ToString();
                        ws.Cells[row, 4].Value = rows["ndtime"] is DBNull ? string.Empty : rows["ndtime"].ToString();
                        //ws.Cells[row, 7].Value = string.IsNullOrEmpty(rows["MachineInterfaceid"].ToString()) ? string.Empty : rows["MachineInterfaceid"].ToString();
                        //ws.Cells[row, 8].Value = string.IsNullOrEmpty(rows["Downinterfaceid"].ToString()) ? string.Empty : rows["Downinterfaceid"].ToString();
                        rowsCount++;
                        row++;
                    }
                    row--;
                    ws.Cells[3, 1, row, 8].AutoFitColumns();
                    ws.Cells[7, 1, row, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[7, 1, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[7, 1, row, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[7, 1, row, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[7, 1, row, 4].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[7, 1, row, 4].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[7, 1, row, 4].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    ws.Cells[7, 1, row, 4].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Calculate();
                }

                Excel.SaveAs(newFile);
                Logger.WriteDebugLog("Downtime Qualification Report Generated Successfully.");

                if (isDataAvailable)
                {
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportFileName);
                    Logger.WriteDebugLog("Downtime Qualification Report Report Exported successfully");
                }
                else
                {
                    Logger.WriteDebugLog("Downtime Qualification Report Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        internal static void ExportJHChecklistTransactionReport(string strReportFile, string ExportPath, string ExportFileName, string Machine, string startTime, string endTime, string plnt, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string RunReportForEvery)
        {
            string dst = string.Empty;
            bool isDataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("JHDashboard Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath,string.Format("JH_Checklist_Transaction_{1}_{0:ddMMMyyyy_HHmmss}.xlsx", DateTime.Parse(startTime),RunReportForEvery));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage Excel = new ExcelPackage(newFile, true);
                System.Data.DataTable dtJHTransaction = AccessReportData.GetJHTransactionDetails(startTime, endTime, Machine);

                if (dtJHTransaction != null && dtJHTransaction.Rows.Count > 0)
                {
                    int rowStart;
                    int colStart = 1;
                    int cntDate = 1, cntShift = 1, cntMachine = 1;
                    isDataAvailable = true;
                    //Excel.Workbook.Worksheets.Delete("Sheet1");

                    var workSheet = Excel.Workbook.Worksheets[1];
                    setWorkSheetSetting(workSheet);
                    rowStart = 4;
                    Machine = Machine.Equals("") ? "ALL" : Machine;
                    workSheet.Cells["A1"].Value = workSheet.Cells["A1"].Value.ToString() + ": " + Machine;
                    //workSheet.Cells["C1"].Value = workSheet.Cells["C1"].Value.ToString() + ": " + "ALL";
                    //workSheet.Cells["F1"].Value = workSheet.Cells["F1"].Value.ToString() + ": " + "ALL";
                    workSheet.Cells["A3"].Value = workSheet.Cells["A3"].Value.ToString() + ": " + DateTime.Parse(startTime).ToString("dd-MMM-yyyy HH:mm:ss");
                    workSheet.Cells["C3"].Value = workSheet.Cells["C3"].Value.ToString() + ": " + DateTime.Parse(endTime).ToString("dd-MMM-yyyy HH:mm:ss");
                    for (int i = 0; i < dtJHTransaction.Rows.Count; i++)
                    {
                        if (i == 0)
                        {
                            workSheet.Cells[rowStart, colStart].Value = "Date";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "Shift";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "Machine ID";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "JH Activity";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "JH Type";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "Status";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            colStart++;
                            workSheet.Cells[rowStart, colStart].Value = "Remarks";
                            workSheet.Cells[rowStart, colStart].Style.Font.Bold = true;
                            workSheet.Cells[rowStart, colStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheet.Cells[rowStart, colStart].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                            rowStart++;
                        }

                        colStart = 1;
                        workSheet.Cells[rowStart, colStart].Value = Convert.ToDateTime(dtJHTransaction.Rows[i]["Sdate"].ToString()).ToString("dd-MMM-yyyy");
                        workSheet.Column(colStart).AutoFit();
                        if (rowStart != 5 && workSheet.Cells[rowStart, colStart].Value.ToString().Equals(workSheet.Cells[rowStart - 1, colStart].Value.ToString()))
                            cntDate++;
                        else
                        {
                            if (rowStart != 5)
                            {
                                workSheet.Cells[rowStart - cntDate, colStart, rowStart - 1, colStart].Merge = true;
                                cntDate = 1;
                            }

                        }
                        colStart++;

                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["ShiftName"].ToString();
                        workSheet.Column(colStart).AutoFit();
                        if (rowStart != 5 && workSheet.Cells[rowStart, colStart].Value.ToString().Equals(workSheet.Cells[rowStart - 1, colStart].Value.ToString()))
                            cntShift++;
                        else
                        {
                            if (rowStart != 5)
                            {
                                workSheet.Cells[rowStart - cntShift, colStart, rowStart - 1, colStart].Merge = true;
                                cntShift = 1;
                            }

                        }
                        colStart++;

                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["Machineid"].ToString();
                        workSheet.Column(colStart).AutoFit();
                        if (rowStart != 5 && workSheet.Cells[rowStart, colStart].Value.ToString().Equals(workSheet.Cells[rowStart - 1, colStart].Value.ToString()))
                            cntMachine++;
                        else
                        {
                            if (rowStart != 5)
                            {
                                workSheet.Cells[rowStart - cntMachine, colStart, rowStart - 1, colStart].Merge = true;
                                cntMachine = 1;
                            }

                        }
                        colStart++;

                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["JHChecklistName"].ToString();
                        workSheet.Column(colStart).AutoFit();
                        colStart++;
                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["JHChecklistType"].ToString();
                        workSheet.Column(colStart).AutoFit();
                        colStart++;
                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["ChecklistStatus"].ToString();
                        workSheet.Column(colStart).AutoFit();
                        colStart++;
                        workSheet.Cells[rowStart, colStart].Value = dtJHTransaction.Rows[i]["Remarks"];
                        workSheet.Column(colStart).AutoFit();
                        colStart++;
                        rowStart++;
                    }
                    rowStart--;
                    if (workSheet.Cells[rowStart, 1].Value.ToString().Equals(workSheet.Cells[rowStart - 1, 1].Value.ToString()))
                    {
                        workSheet.Cells[rowStart - cntDate + 1, 1, rowStart, 1].Merge = true;
                        cntDate = 1;
                    }

                    if (workSheet.Cells[rowStart, 2].Value.ToString().Equals(workSheet.Cells[rowStart - 1, 2].Value.ToString()))
                    {
                        workSheet.Cells[rowStart - cntShift + 1, 2, rowStart, 2].Merge = true;
                        cntShift = 1;
                    }

                    if (workSheet.Cells[rowStart, 3].Value.ToString().Equals(workSheet.Cells[rowStart - 1, 3].Value.ToString()))
                    {
                        workSheet.Cells[rowStart - cntMachine + 1, 3, rowStart, 3].Merge = true;
                        cntMachine = 1;
                    }

                    workSheet.Cells[4, 1, rowStart, colStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    workSheet.Cells[4, 1, rowStart, colStart].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells[4, 1, rowStart, 7].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);

                    
                }
                Excel.SaveAs(newFile);
                Logger.WriteDebugLog("JH Transaction Report has been generated successfully.");

                if (isDataAvailable)
                {
                    Logger.WriteDebugLog("JH Transaction Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportFileName);
                }
                else
                {
                    Logger.WriteDebugLog("JH Transaction Report not mailed: no data");
                }
            }
            catch(Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
                throw;
            }
        }

        internal static void setWorkSheetSetting(ExcelWorksheet wksheet)
        {
            wksheet.PrinterSettings.Orientation = eOrientation.Landscape;
            wksheet.PrinterSettings.FitToPage = true;
            wksheet.PrinterSettings.FitToWidth = 1;
            wksheet.PrinterSettings.FitToHeight = 0;

        }
        public static bool ExportShiftProductionCountHourly(string strReportFile, string ExportPath, string ExportedReportFile,
           int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
           string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string CompanyName, bool MachineAE) //vasavi do to
        {

            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet wrksht = null;
            object misValue = System.Reflection.Missing.Value;
            int pid = 0;

            try
            {
                string src, dst = string.Empty;//Globally Used  
                string plantname = plantid;
                string SDate = string.Format("{0:yyyy-MMM-dd hh:mm:ss tt}", sttime);
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\ShiftProductionCountHourlyBoschBNGTemplate2.xlsx";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }

                plantname = (plantid == "") ? "All Plant" : plantid;
                int indx_col = AccessReportData.MaxHourIdShift();

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = ExportPath + @"ShiftProductionCountHourlyBoschBNG2_" + plantname + "_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(sttime).AddDays(DayBefores)) + ".xlsx";//string.Format("{0:hh-mm-ss MMM-yyyy}", DT) + ".xls";
                try
                {
                    File.Copy(src, dst, true);
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }

                if (!File.Exists(dst))
                {
                    return false;
                }

                //fetch all the machines, create a worksheet for all the machines
                List<string> machineList = default(List<string>);
                if ((string.IsNullOrEmpty(MachineId) || MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase)) && (!string.IsNullOrEmpty(plantname)))
                {
                    machineList = AccessReportData.GetTPMTrakEnabledMachines(plantname);
                }

                else
                {
                    machineList = new List<string>();
                    machineList.Add(MachineId);
                }

                //for each worksheet do the below things
                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;

                int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);

                for (int i = 1; i < machineList.Count; i++)
                {
                    wrksht.Copy(Missing.Value, wrksht);
                }
                Logger.WriteDebugLog("Generating reports for " + DateTime.Parse(sttime).AddDays(DayBefores).ToString());
                for (int i = 0; i < machineList.Count; i++)
                {
                    string machineName = machineList[i];
                    wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(i + 1);
                    wrksht.Name = machineName;

                    int Row = 16;
                    int col = 2;

                    wrksht.Cells[11, 1] = "Hourly Tracking - " + machineName;
                    //    wrksht.Cells[11, 43] = string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(sttime).AddDays(DayBefores));
                    wrksht.Cells[11, 44] = string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(sttime).AddDays(DayBefores));

                    SqlDataReader rs = AccessReportData.ShiftProductionCountHour(DateTime.Parse(sttime).AddDays(DayBefores), machineName, "BOSCH_BNG_CamShaft");

                    while (rs.Read())
                    {
                        wrksht.Cells[Row, col] = rs["Target"];
                        wrksht.Cells[Row, col + 2] = rs["Actual"];
                        wrksht.Cells[Row, 34] = rs["KWH"];
                        Row = Row + 1;
                        if (Convert.ToInt32(rs["HourID"]) == 8)
                        {
                            Row = Row + 1;
                        }
                    }
                    if (rs != null)
                    {
                        rs.Close();
                    }

                    Row = 16;
                    col = 37;
                    int prevHour, prevShift;
                    prevHour = 1; prevShift = 1;
                    rs = AccessReportData.ShiftProductionCountHour(DateTime.Parse(sttime).AddDays(DayBefores), machineName, "BOSCH_BNG_AELosses");
                    while (rs.Read())
                    {
                        if (int.Parse(rs["HourID"].ToString()) == prevHour)
                        {
                            col++;
                        }
                        else
                        {
                            col = 38;
                            Row++;
                        }
                        Row = (prevShift != int.Parse(rs["Shiftid"].ToString())) ? Row + 1 : Row;
                        prevShift = int.Parse(rs["Shiftid"].ToString());

                        if (int.Parse(rs["HourID"].ToString()) == 1 && int.Parse(rs["Shiftid"].ToString()) == 1)
                        {
                            wrksht.Cells[Row - 1, col] = rs["DownCategory"];
                        }

                        wrksht.Cells[Row, col] = rs["DownTime"];
                        prevHour = int.Parse(rs["HourID"].ToString());
                    }

                    if (rs != null)
                    {
                        rs.Close();
                    }
                }
                ////run the macro
                //xlApp.Run("plotgraph", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value
                //    , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value
                //    , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                Excel.Worksheet xlWorkSheetFocus = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                xlWorkSheetFocus.Activate();
                if (wrkbk != null) wrkbk.Close(true, misValue, misValue);
                if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                if (xlApp != null)
                {
                    releaseObject(wrksht);
                    releaseObject(wrkbk);
                    releaseObject(xlApp);
                }
                if (pid != 0) KillSpecificExcelFileProcess(pid);
            }
            return true;
        }


        public static bool ExportHorlypartsCountReportFormatI(string strReportFile, string ExportPath, string ExportedReportFile,
         int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
         string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
         string Email_List_BCC, string CompanyName, bool MachineAE) //vasavi do to
        {

            //Excel.Application xlApp = null;
            //Excel.Workbook wrkbk = null;
            //Excel.Worksheet wrksht = null;
            //object misValue = System.Reflection.Missing.Value;
            //FileInfo newFile = new FileInfo(dst);
            //ExcelPackage excelPackage = new ExcelPackage(newFile, true);
            //ExcelWorksheet wrksht = excelPackage.Workbook.Worksheets[1];
            int pid = 0;

            try
            {
                string src, dst = string.Empty;//Globally Used  
                string plantname = plantid;
                string SDate = string.Format("{0:yyyy-MMM-dd hh:mm:ss tt}", sttime);
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\HourlyMonitoringReportBng.xlsx";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }

                plantname = (plantid == "") ? "All Plant" : plantid;
                int indx_col = AccessReportData.MaxHourIdShift();

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = ExportPath + @"HourlyMonitoringReportBng_" + plantname + "_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(sttime).AddDays(DayBefores)) + ".xlsx";//string.Format("{0:hh-mm-ss MMM-yyyy}", DT) + ".xls";
                try
                {
                    File.Copy(src, dst, true);
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }

                if (!File.Exists(dst))
                {
                    return false;
                }

                //fetch all the machines, create a worksheet for all the machines
                List<string> machineList = default(List<string>);
                if ((string.IsNullOrEmpty(MachineId) || MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase)) && (!string.IsNullOrEmpty(plantname)))
                {
                    machineList = AccessReportData.GetTPMTrakEnabledMachines(plantname);
                }

                else
                {
                    machineList = new List<string>();
                    machineList.Add(MachineId);
                }

                ////for each worksheet do the below things
                //xlApp = new Excel.ApplicationClass();
                //xlApp.DisplayAlerts = false;

                //int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                //wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                // var templateSheet = excelPackage.Workbook.Worksheets[1];

                //for (int i = 1; i < machineList.Count; i++)
                //{
                //  excelPackage.Workbook.Worksheets.Copy(
                //}
                //   var someName = excelPackage.Workbook.Worksheets.Add("someName", templateSheet);
                Logger.WriteDebugLog("Generating reports for " + DateTime.Parse(sttime).AddDays(DayBefores).ToString());
                for (int i = 0; i < machineList.Count; i++)
                {
                    string machineName = machineList[i];
                    ExcelWorksheet wrksht = excelPackage.Workbook.Worksheets[1];
                    //  wrksht.Name = machineName;

                    int Row = 21;
                    int col = 2;

                    wrksht.Cells[11, 1].Value = "Hourly Tracking - " + machineName;
                    wrksht.Cells[11, 43].Value = string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(sttime).AddDays(DayBefores));

                    SqlDataReader rs = AccessReportData.ShiftProductionCountHour(DateTime.Parse(sttime).AddDays(DayBefores), machineName, "BOSCH_BNG_CamShaft");

                    while (rs.Read())
                    {
                        wrksht.Cells[Row, 2].Value = rs["Target"];
                        wrksht.Cells[Row, 7].Value = rs["Actual"];
                        Row = Row + 1;
                        if (Convert.ToInt32(rs["HourID"]) == 8)
                        {
                            Row = Row + 1;
                        }
                        // = ((int)rs["HourID"] == 8) ? Row + 1 : Row;
                        //Row++;
                    }
                    if (rs != null)
                    {
                        rs.Close();
                    }

                }

                //Excel.Worksheet xlWorkSheetFocus = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                // xlWorkSheetFocus.Activate();

                excelPackage.SaveAs(newFile);

                //if (wrkbk != null)

                //    wrkbk.Close(true, misValue, misValue);
                //if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                //if (xlApp != null)
                //{
                //    releaseObject(wrksht);
                //    releaseObject(wrkbk);
                //    releaseObject(xlApp);
                //}
                if (pid != 0) KillSpecificExcelFileProcess(pid);
            }
            return true;
        }


        public static bool ExportOEETrend(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
        string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
        string Email_List_BCC, string CompanyName, bool MachineAE)
        {
            string src, dst = string.Empty;//Globally Used
            int pid = 0;
            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet wrksht = null;
            Excel.Range workSheet_range = null;
            try
            {
                SqlConnection Con = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand();
                SqlDataReader reader;

                string plantname = plantid;
                string machinename = MachineId;
                string shiftname = Shift;
                string strsql = string.Empty;
                int rowcount = 0;
                int rowposit = 0;
                string currmonth, currmonth1;

                string SDate = string.Format("{0:yyyy-MMM-dd hh:mm:ss tt}", sttime);
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\OEE Trend.xls";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"OEETrend_" + plantname + "_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(sttime)) + ".xls";

                if (!File.Exists(dst))
                {
                    File.Copy(src, dst, true);
                }

                strsql = "select distinct M.Machineid,P.Plantid,0 from Machineinformation M inner join PlantMachine P on M.machineid=P.Machineid inner join shiftproductiondetails SPD on SPD.machineid=M.machineid where M.TPMTrakEnabled='1' ";

                if (plantname != string.Empty)
                    strsql = strsql + " and P.PlantID='" + plantname + "' ";
                if (machinename != string.Empty)
                    strsql = strsql + " and M.Machineid='" + machinename + "' ";
                if (machinename != string.Empty)
                    machinename = "'" + machinename + "'";

                cmd = new SqlCommand(strsql, Con);
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        rowcount++;
                    }
                }
                Con.Close();
                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "All Machines", "Format1");

                object misValue = System.Reflection.Missing.Value;

                if (!File.Exists(dst))
                {
                    return false;
                }
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);

                wrksht.Cells[1, 2] = "ONLINE TRACKING OF OEE TREND ";
                wrksht.Cells[2, 5] = string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(sttime));
                wrksht.Cells[2, 8] = string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(ndtime));
                if (shiftname != string.Empty)
                    wrksht.Cells[2, 14] = shiftname;
                else
                    wrksht.Cells[2, 14] = "ALL";

                wrksht.Cells[3, 4] = "" + string.Format("{0: yyyy}", DateTime.Parse(sttime).AddYears(-1)) + " - " + (string.Format("{0:yyyy}", DateTime.Parse(sttime))) + "";
                wrksht.Cells[3, 17] = "" + string.Format("{0:yyyy}", DateTime.Parse(sttime)) + " - " + string.Format("{0:yyyy}", DateTime.Parse(sttime).AddYears(1)) + "";

                int Row = 4;
                int col = 1;
                int i = 0;

                for (i = 0; i < rowcount; i++)
                {
                    reader.Read();
                    col = 1;
                    wrksht.Cells[Row, col] = reader["Plantid"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["MachineID"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["ownername"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["prevyearoee"].ToString();

                    col = 17;
                    wrksht.Cells[Row, col] = reader["machinewisetarget"].ToString();
                    Row++;
                }

                /* ------------------------------------------------------------------ */
                /* For displaying the data for Month wise  */
                /* Starts here------------------------------------------------------- */
                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "All Machines", "Format1");
                Row = 4;
                col = 5;

                if (reader.Read())
                {
                    while (true)
                    {
                        currmonth = reader["Pdate"].ToString();
                        wrksht.Cells[3, col] = reader["Pdate"].ToString();
                        wrksht.Cells[Row, col] = reader["oeffy"].ToString();
                        Row = Row + 1;

                        if (reader.Read())
                        {
                            currmonth1 = reader["Pdate"].ToString();
                            if (currmonth1 != currmonth)
                            {
                                currmonth = currmonth1;
                                col = col + 1;
                                wrksht.Cells[3, col] = reader["Pdate"].ToString();
                                Row = 4;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                reader.Close();

                workSheet_range = wrksht.get_Range("D3", "D" + (Row - 1));
                workSheet_range.Interior.ColorIndex = 43;
                workSheet_range = wrksht.get_Range("Q3", "Q" + (Row - 1));
                workSheet_range.Interior.ColorIndex = 43;
                workSheet_range = wrksht.get_Range("A3", "Q" + (Row - 1));
                workSheet_range.Borders.ColorIndex = 1;
                workSheet_range = wrksht.get_Range("C3", "Q" + (Row - 1));
                workSheet_range.HorizontalAlignment = -4108;
                workSheet_range = wrksht.get_Range("C3", "Q" + (Row - 1));
                workSheet_range.VerticalAlignment = -4108;

                /* Ends here----------------------------------------------------------- */

                #region /* ----------------------Format 2 ------------------------------------- */
                /* Starts here--------------------------------------------------------- */


                strsql = "select Count(machineid) as mcount from machineinformation where tpmtrakenabled=1 and CriticalMachineenabled=1";
                rowposit = Row;
                Row = Row + 3;
                col = 1;
                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "Critical Machines", "Format2");


                wrksht.Cells[Row, col] = "OEE FOR FOCUSED MACHINE";
                workSheet_range = wrksht.get_Range("A" + Row, "A" + Row);
                workSheet_range.Interior.ColorIndex = 40;
                workSheet_range = wrksht.get_Range("A" + Row, "A" + Row);
                workSheet_range.HorizontalAlignment = -4108;
                workSheet_range = wrksht.get_Range("A" + Row, "A" + Row);
                workSheet_range.VerticalAlignment = -4108;
                workSheet_range = wrksht.get_Range("A" + Row, "A" + Row);
                try
                {
                    workSheet_range = wrksht.get_Range("A" + Row, "E" + Row);
                    workSheet_range.Merge(Missing.Value);
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                workSheet_range = wrksht.get_Range("A" + Row, "E" + Row);
                workSheet_range.Font.Bold = true;
                workSheet_range = wrksht.get_Range("A" + Row, "E" + Row);
                workSheet_range.HorizontalAlignment = -4108;
                workSheet_range = wrksht.get_Range("A" + Row, "E" + Row);
                workSheet_range.VerticalAlignment = -4108;

                wrksht.Cells[Row + 1, 1] = "Month";
                workSheet_range = wrksht.get_Range("A" + (Row + 1), "A1");
                workSheet_range.HorizontalAlignment = -4108;
                workSheet_range = wrksht.get_Range("A" + (Row + 1), "A1");
                workSheet_range.VerticalAlignment = -4108;
                wrksht.Cells[Row + 1, 2] = "OEE";
                wrksht.Cells[Row + 1, 3] = "A";
                wrksht.Cells[Row + 1, 4] = "P";
                wrksht.Cells[Row + 1, 5] = "QR";

                workSheet_range = wrksht.get_Range("A" + (Row + 1), "E" + (Row + 1));
                workSheet_range.Interior.ColorIndex = 37;
                workSheet_range = wrksht.get_Range("A" + (Row + 1), "E" + (Row + 1));
                workSheet_range.Font.Bold = true;
                int rwcol = Row + 2;
                reader.Read();
                wrksht.Cells[Row + 2, 1] = string.Format("{0:yyyy}", DateTime.Parse(sttime).AddYears(-1)) + " - " + string.Format("{0:yyyy}", DateTime.Parse(sttime));
                wrksht.Cells[Row + 2, 2] = reader["prevyearoee"].ToString();
                wrksht.Cells[Row + 14, 1] = "Target";
                wrksht.Cells[Row + 14, 2] = reader["machinewisetarget"].ToString();
                Row = Row + 3;

                while (reader.Read())
                {
                    col = 1;
                    wrksht.Cells[Row, col] = reader["Pdate"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["oeffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["aeffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["peffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["qeffy"].ToString();

                    Row = Row + 1;
                }
                reader.Close();

                //wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);

                workSheet_range = wrksht.get_Range("A" + rwcol, "A" + Row);
                workSheet_range.Font.ColorIndex = 30;
                workSheet_range = wrksht.get_Range("A" + rwcol, "A" + Row);
                workSheet_range.Font.Bold = true;
                workSheet_range = wrksht.get_Range("A" + (rowposit + 3), "E" + Row);
                workSheet_range.Borders.Color = 1;
                workSheet_range = wrksht.get_Range("B" + Row, "B" + Row);
                workSheet_range.Interior.ColorIndex = 43;
                workSheet_range = wrksht.get_Range("B" + rwcol, "B" + rwcol);
                workSheet_range.Interior.ColorIndex = 53;

                workSheet_range = wrksht.get_Range("B" + (rowposit + 3), "E" + Row);
                workSheet_range.HorizontalAlignment = -4108;
                workSheet_range = wrksht.get_Range("B" + (rowposit + 3), "E" + Row);
                workSheet_range.VerticalAlignment = -4108;
                workSheet_range = wrksht.get_Range("A" + (rwcol - 2), "E" + Row);
                workSheet_range.EntireColumn.AutoFit();

                #endregion /* Ends here--------------------------------------------------------- */

                #region /* ----------------------Format 3 ------------------------------------- */
                /* Starts here--------------------------------------------------------- */
                Row = rowposit;
                Row = Row + 3;
                col = 7;

                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "ALL Machines", "Format3");

                wrksht.Cells[Row, col] = "OVERALL PLANT OEE";
                int rowval = Row;
                workSheet_range = wrksht.get_Range("G" + Row, "G" + Row);
                workSheet_range.Interior.ColorIndex = 40;
                workSheet_range = wrksht.get_Range("G" + Row, "K" + Row);
                try
                {
                    workSheet_range.Merge(Missing.Value);
                    workSheet_range = wrksht.get_Range("G" + Row, "K" + Row);
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                workSheet_range.Font.Bold = true;
                wrksht.Cells[Row + 1, 7] = "Month";
                wrksht.Cells[Row + 1, 8] = "OEE";
                wrksht.Cells[Row + 1, 9] = "A";
                wrksht.Cells[Row + 1, 10] = "P";
                wrksht.Cells[Row + 1, 11] = "QR";
                workSheet_range = wrksht.get_Range("G" + (Row + 1), "K" + (Row + 1));
                workSheet_range.Interior.ColorIndex = 37;
                workSheet_range = wrksht.get_Range("G" + (Row + 1), "K" + (Row + 1));
                workSheet_range.Font.Bold = true;
                wrksht.Cells[Row + 13, 7] = "Target";
                reader.Read();
                wrksht.Cells[Row + 13, 8] = reader["machinewisetarget"].ToString();
                Row = Row + 2;

                while (reader.Read())
                {
                    col = 7;
                    wrksht.Cells[Row, col] = reader["Pdate"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["oeffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["aeffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["peffy"].ToString();

                    col = col + 1;
                    wrksht.Cells[Row, col] = reader["qeffy"].ToString();

                    Row = Row + 1;
                }

                reader.Close();

                workSheet_range = wrksht.get_Range("G" + rowval, "G" + Row);
                workSheet_range.Font.ColorIndex = 30;
                workSheet_range = wrksht.get_Range("G" + rowval, "G" + Row);
                workSheet_range.Font.Bold = true;
                workSheet_range = wrksht.get_Range("G" + rowval, "k" + Row);
                workSheet_range.Borders.Color = 1;
                workSheet_range = wrksht.get_Range("H" + Row, "H" + Row);
                workSheet_range.Interior.ColorIndex = 43;

                #endregion /* Ends here--------------------------------------------------------- */

                #region  /* ----------------------Format 4 ------------------------------------- */
                /* Starts here--------------------------------------------------------- */

                Excel.Worksheet wrksht1 = null;

                int sheetno, C_downid;
                int row1, col1;
                string Cur_machine = string.Empty;
                sheetno = 2;
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(sheetno);

                strsql = "Select Count(machineid) as mcount from machineinformation where tpmtrakenabled=1 ";
                cmd = new SqlCommand(strsql, Con);
                Con.Open();
                rowcount = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                Con.Close();

                strsql = "select count(downid) as downid from downcodeinformation ";
                cmd = new SqlCommand(strsql, Con);
                Con.Open();
                C_downid = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                Con.Close();

                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "ALL Machines", "Format4");
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(sheetno);
                Row = 1;
                col = 1;
                if (sttime == string.Empty && ndtime == string.Empty)
                {
                    wrksht.Cells[4, 1] = string.Empty;
                }
                else
                {
                    wrksht.Cells[4, 1] = string.Format("{0:yyyy}", DateTime.Parse(sttime).AddYears(-1)) + " - " + string.Format("{0:yyyy}", DateTime.Parse(sttime).AddYears(-1));
                }

                if (reader.HasRows)
                {
                    reader.Read();

                    if (reader["prevyearoee"].ToString() == string.Empty)
                    {
                        wrksht.Cells[4, 2] = string.Empty;
                    }
                    else
                    {
                        wrksht.Cells[4, 2] = reader["prevyearoee"].ToString();
                    }
                    //wrksht.Cells[17, 1] = "Target";

                    if (reader["machinewisetarget"].ToString() == string.Empty)
                    {
                        wrksht.Cells[17, 2] = string.Empty;
                    }
                    else
                    {
                        wrksht.Cells[17, 2] = reader["machinewisetarget"].ToString();
                    }
                    wrksht.Cells[2, 2] = "" + string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(sttime));
                    wrksht.Cells[2, 5] = "" + string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(ndtime));

                    if (shiftname != string.Empty)
                        wrksht.Cells[2, 6] = "Shift: " + shiftname + " ";
                    else
                        wrksht.Cells[2, 6] = " Shift: ALL";

                    Cur_machine = string.Empty;
                    Row = 3;

                    while (reader.Read())
                    {
                        if (Cur_machine == reader["machineid"].ToString())
                        {
                            wrksht.Cells[1, 1] = Cur_machine;
                            col = 1;
                            col1 = 1;
                            wrksht.Cells[Row, col] = reader["Pdate"].ToString();

                            col = col + 2;
                            wrksht.Cells[Row, col] = reader["oeffy"].ToString();

                            col = col + 1;
                            wrksht.Cells[Row, col] = reader["aeffy"].ToString();

                            col = col + 1;
                            wrksht.Cells[Row, col] = reader["peffy"].ToString();

                            col = col + 1;
                            wrksht.Cells[Row, col] = reader["qeffy"].ToString();
                            wrksht.Cells[4, 1] = "" + string.Format("{0:yyyy}", DateTime.Parse(sttime).AddYears(-1)) + " - " + string.Format("{0:yyyy}", DateTime.Parse(sttime));
                            wrksht.Cells[4, 2] = reader["prevyearoee"].ToString();
                            wrksht.Cells[17, 1] = "Target";
                            wrksht.Cells[17, 2] = reader["machinewisetarget"].ToString();
                            Row = Row + 1;
                        }

                        else
                        {
                            if (sheetno > 2)
                                wrksht.Copy(Missing.Value, (Excel.Worksheet)wrkbk.Worksheets.get_Item(sheetno - 1));

                            string Shtname;
                            Shtname = wrksht.Name.ToString();

                            reader.Read();
                            wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(sheetno);
                            // Shtname = reader["MachineID"].ToString();
                            if (Shtname == reader["MachineID"].ToString())
                                wrksht.Name = reader["MachineID"].ToString() + " ";
                            else
                                wrksht.Name = reader["MachineID"].ToString();

                            Cur_machine = wrksht.Name;
                            Row = 5;
                            col = 1;
                            Cur_machine = reader["machineid"].ToString();
                            wrksht.Columns.AutoFit();

                            Excel.ChartObjects chartObjects = (Excel.ChartObjects)wrksht.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)chartObjects.Item(1);
                            myChart.Chart.HasTitle = true;
                            myChart.Chart.ChartTitle.Text = Cur_machine;

                            myChart = (Excel.ChartObject)chartObjects.Item(2);
                            myChart.Chart.HasTitle = true;
                            myChart.Chart.ChartTitle.Text = Cur_machine + " PARETO LOSS ";

                            sheetno++;
                        }
                    }
                }
                reader.Close();
                #endregion  /* Ends here--------------------------------------------------------- */

                #region  /* ----------------------Format 5 ------------------------------------- */
                /* Starts here--------------------------------------------------------- */
                string check = string.Empty;
                row1 = 23;
                col1 = 1;
                sheetno = 2;
                reader = AccessReportData.OEETrend(string.Format("{0:yyyy-MM-dd}", DateTime.Parse(sttime)), string.Format("{0:yyyy-MM-dd}", DateTime.Parse(ndtime)), shiftname, plantname, machinename, "ALL Machines", "Format5");

                while (reader.Read())
                {
                    //wrksht.Cells[1, 1] = wrksht.Name;

                    if (reader["machineid"].ToString() == Cur_machine)
                    {
                        col1 = 1;
                        // wrksht.Cells[row1, col1] = reader["downid"].ToString();

                        col1 = col1 + 1;
                        wrksht.Cells[row1, col1] = reader["downtime"].ToString();

                        row1 = row1 + 1;
                        wrksht.Cells[1, 2] = reader["machineid"].ToString();
                        //rs_format5.MoveNext
                    }
                    else
                    {
                        wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(sheetno);
                        wrksht.Cells[1, 2] = reader["machineid"].ToString();
                        //wrksht.Cells[1, 1] = reader["machineid"].ToString();
                        row1 = 23;
                        col1 = 1;
                        Cur_machine = reader["machineid"].ToString();
                    }

                    wrksht.Cells[row1, 1] = reader["downid"].ToString();

                    wrksht.Cells[23, 3] = wrksht.Cells[23, 2];

                    workSheet_range = wrksht.get_Range("B23", "B62");
                    ((Range)workSheet_range.Cells[20, 2]).Value2 = "=Sum(B23:B62)";
                }

                /* Sets the default sheet to Sheet-1 */
                /* ---------------------------------------- */
                Excel.Worksheet xlWorkSheetFocus = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                xlWorkSheetFocus.Activate();
                /* ---------------------------------------- */
                wrkbk.Close(true, misValue, misValue);
                if (xlApp != null) xlApp.Quit();
                reader.Close();
                #endregion

                Logger.WriteDebugLog("OEE Report Generated Successfully.");
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }

            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Report generation Failed. Error:{0}.", ex.ToString()));
                return false;
            }
            finally
            {
                if (xlApp != null)
                {
                    releaseObject(wrksht);
                    releaseObject(wrkbk);
                    releaseObject(xlApp);
                }
                if (pid != 0) KillSpecificExcelFileProcess(pid);
            }
            return true;
        }

        private static void KillSpecificExcelFileProcess(int pid)
        {
            try
            {
                Process processes = null;
                processes = Process.GetProcessById(pid);
                if (processes != null) processes.Kill();
            }
            catch (Exception ex)
            {
            }
        }

        public static void WriteInToFile(string Location, string str)
        {
            StreamWriter writer = new StreamWriter(Location, true, Encoding.Default, 8195);
            writer.WriteLine(str);
            str = string.Empty;
            writer.Flush();
            writer.Close();
            writer.Dispose();
        }

        private static void releaseObject(Object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static string GetZipFile(string dst)
        {
            string filePath = dst;
            string sourceFolder = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileName(filePath);
            string ZipFileName = Path.Combine(sourceFolder, fileName.Replace(".xls", ".zip"));

            try
            {
                if (File.Exists(ZipFileName))
                {
                    File.Delete(ZipFileName);
                }

                FileStream fsOut = File.Create(ZipFileName);
                ZipOutputStream zipStream = new ZipOutputStream(fsOut);
                zipStream.SetLevel(9); //0-9, 9 being the highest level of compression
                FileInfo fi = new FileInfo(filePath);
                ZipEntry newEntry = new ZipEntry(fileName);
                newEntry.DateTime = fi.LastWriteTime; // Note the zip format stores 2 second granularity
                newEntry.Size = fi.Length;
                zipStream.PutNextEntry(newEntry);
                byte[] buffer = new byte[4096];
                using (FileStream streamReader = File.OpenRead(filePath))
                {
                    StreamUtils.Copy(streamReader, zipStream, buffer);
                }
                zipStream.CloseEntry();

                zipStream.IsStreamOwner = true; // Makes the Close also Close the underlying stream
                zipStream.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(string.Format("Exception : Zipping of the files {0}", ex.ToString()));
            }
            return ZipFileName;
        }

        public static bool ExportDailyProductionReportTrellBorg(string strReportFile, string ExportPath, string ExportedReportFile,
                    int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
                    string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                    string Email_List_BCC, string CompanyName, bool MachineAE, string componentId, string operationId)
        {

            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet Wrksht_down;
            Excel.Worksheet wrksht = null;
            object misValue = System.Reflection.Missing.Value;
            int pid = 0;
            try
            {

                string src, dst = string.Empty;

                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\SM_DailyProductionReport.xls";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }
                plantid = (plantid.ToUpper() == "ALL") ? "" : plantid;
                MachineId = (MachineId.ToUpper() == "ALL") ? "" : MachineId;

                SqlDataReader DR = AccessReportData.GetExportReports(string.Empty);
                if (DR.HasRows)
                {
                    DR.Read();
                    ExportPath = Convert.ToString(DR["ExportPath"]);
                }
                else
                {
                    ExportPath = APath + @"\Reports\Temp\";
                }
                if (DR != null) DR.Close();

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = ExportPath + @"SM_DailyProductionReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                File.Copy(src, dst, true);

                Thread.Sleep(1000);
                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;

                int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);

                if (!File.Exists(dst))
                {
                    return false;
                }

                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                string name = wrksht.Name;
                SqlDataReader rs = AccessReportData.ProductionTrelBorg(DateTime.Parse(sttime), DateTime.Parse(ndtime), plantid, MachineId, string.Empty, string.Empty, 0);
                int row = 4, interior = 1, row1 = 0;
                string plantName = string.Empty;
                string machineName = string.Empty;
                string dDate = string.Empty;
                wrksht.Cells[1, 1] = CompanyName;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        if (!Convert.IsDBNull(rs["Day"]))
                        {
                            wrksht.Cells[row, 1] = Convert.ToDateTime(rs["Day"]).ToString("dd-MM-yyyy");
                        }

                        if ((machineName != Convert.ToString(rs["Machine"]) || machineName == string.Empty) || (dDate == string.Empty || dDate != Convert.ToString(rs["Day"])))
                        {
                            dDate = Convert.ToString(rs["Day"]);
                            row1 = row;
                            wrksht.Cells[row, 2] = Convert.ToString(rs["Cell"]);
                            plantName = Convert.ToString(rs["Cell"]);
                            wrksht.Cells[row, 3] = Convert.ToString(rs["Machine"]);
                            machineName = Convert.ToString(rs["Machine"]);
                            wrksht.Cells[row, 13] = Convert.ToString(rs["ProdTime"]);
                            wrksht.Cells[row, 14] = Convert.ToString(rs["DownTime"]);
                            wrksht.Cells[row, 15] = Convert.ToString(rs["SettingTime"]);
                            wrksht.Cells[row, 16] = Convert.ToString(rs["OverallEfficiency"]);
                            wrksht.Cells[row, 17] = Convert.ToString(rs["DownReason1"]);
                            wrksht.Cells[row, 18] = Convert.ToString(rs["DownReason2"]);
                            wrksht.Cells[row, 19] = Convert.ToString(rs["Operator"]);
                            interior = interior + 1;
                        }

                        //Excel.Range c1 = (Excel.Range)wrksht.Cells[row1, 2];
                        //Excel.Range c2 = (Excel.Range)wrksht.Cells[row,2];
                        //Excel.Range range = wrksht.get_Range(c1, c2);
                        //range.Merge();

                        //Excel.Range c3 = (Excel.Range)wrksht.Cells[row1, 3];
                        //Excel.Range c4 = (Excel.Range)wrksht.Cells[row, 3];
                        //Excel.Range range1 = wrksht.get_Range(c3, c4);
                        //range1.Merge();

                        wrksht.get_Range(wrksht.Cells[row1, 2], wrksht.Cells[row, 2]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 3], wrksht.Cells[row, 3]).Merge(false);

                        wrksht.Cells[row, 4] = Convert.ToString(rs["Component"]);
                        wrksht.Cells[row, 5] = Convert.ToString(rs["RunTime"]);
                        wrksht.Cells[row, 6] = Convert.ToString(rs["Target"]);
                        wrksht.Cells[row, 7] = Convert.ToString(rs["CountShift1"]);
                        wrksht.Cells[row, 8] = Convert.ToString(rs["CountShift2"]);
                        wrksht.Cells[row, 9] = Convert.ToString(Convert.ToUInt32(rs["CountShift1"]) + Convert.ToUInt32(rs["CountShift2"]));
                        wrksht.Cells[row, 10] = Convert.ToString(rs["frmtStdCycletime"]);
                        wrksht.Cells[row, 11] = Convert.ToString(rs["frmtActualCycletime"]);
                        wrksht.Cells[row, 12] = Convert.ToString(rs["Cyclefficiency"]);
                        wrksht.get_Range(wrksht.Cells[row1, 13], wrksht.Cells[row, 13]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 14], wrksht.Cells[row, 14]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 15], wrksht.Cells[row, 15]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 16], wrksht.Cells[row, 16]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 17], wrksht.Cells[row, 17]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 18], wrksht.Cells[row, 18]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 19], wrksht.Cells[row, 19]).Merge(false);
                        if (interior % 2 == 0)
                        {
                            wrksht.get_Range(wrksht.Cells[row, 1], wrksht.Cells[row1, 19]).Interior.ColorIndex = 40;
                        }
                        row++;
                    }
                }
                else
                {
                    Logger.WriteDebugLog("No records found.");
                }

                if (rs != null)
                {
                    rs.Close();
                }
                wrksht.Columns.AutoFit();
                wrksht.ShowAllData();
                //wrkbk.Close(true, misValue, misValue);
                wrkbk.SaveAs(dst, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange,
                        false, false, Missing.Value, Missing.Value, Missing.Value);
                if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");
                Thread.Sleep(2000);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                if (xlApp != null)
                {
                    releaseObject(wrkbk);
                    releaseObject(wrksht);
                    releaseObject(xlApp);
                }
                if (pid != 0)
                    KillSpecificExcelFileProcess(pid);

            }
            return true;
        }


        public static bool ExportDailyProductionReportTrellBorgDayWise(string strReportFile, string ExportPath, string ExportedReportFile,
                  int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
                  string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                  string Email_List_BCC, string CompanyName, bool MachineAE, string componentId, string operationId)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet wrksht = null;
            int pid = 0;
            object misValue = System.Reflection.Missing.Value;

            //sttime = "2018-Jan-20 06:00:00 AM"; //g:test
            //ndtime = "2018-Jan-21 06:00:00 AM";
            try
            {
                string src, dst = string.Empty;
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\SM_DailyProductionReport.xls";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }
                plantid = (plantid.ToUpper() == "ALL") ? "" : plantid;
                MachineId = (MachineId.ToUpper() == "ALL") ? "" : MachineId;

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = ExportPath + @"SM_DailyProductionReport_" + DateTime.Parse(sttime).ToString("yyyyMMdd") + ".xls";
                if (File.Exists(dst))
                {
                    File.Delete(dst);
                }
                File.Copy(src, dst, true);

                Thread.Sleep(1000);
                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;
                int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                if (!File.Exists(dst))
                {
                    return false;
                }
                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                string name = wrksht.Name;
                SqlDataReader rs = AccessReportData.ProductionTrelBorg(DateTime.Parse(sttime), DateTime.Parse(sttime), plantid, MachineId, string.Empty, string.Empty, DayBefores);
                int row = 4, interior = 1, row1 = 0;
                string plantName = string.Empty;
                string machineName = string.Empty;
                string dDate = string.Empty;
                wrksht.Cells[1, 1] = CompanyName;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        if (!Convert.IsDBNull(rs["Day"]))
                        {
                            wrksht.Cells[row, 1] = Convert.ToDateTime(rs["Day"]).ToString("dd-MM-yyyy");
                        }

                        if ((machineName != Convert.ToString(rs["Machine"]) || machineName == string.Empty) || (dDate == string.Empty || dDate != Convert.ToString(rs["Day"])))
                        {
                            dDate = Convert.ToString(rs["Day"]);
                            row1 = row;
                            wrksht.Cells[row, 2] = Convert.ToString(rs["Cell"]);
                            plantName = Convert.ToString(rs["Cell"]);
                            wrksht.Cells[row, 3] = Convert.ToString(rs["Machine"]);
                            machineName = Convert.ToString(rs["Machine"]);

                            wrksht.Cells[row, 14] = Convert.ToString(rs["ProdTime"]);
                            wrksht.Cells[row, 15] = Convert.ToString(rs["DownTime"]);
                            wrksht.Cells[row, 16] = Convert.ToString(rs["SettingTime"]);
                            wrksht.Cells[row, 17] = Convert.ToString(rs["OverallEfficiency"]);
                            wrksht.Cells[row, 18] = Convert.ToString(rs["DownReason1"]);
                            wrksht.Cells[row, 19] = Convert.ToString(rs["DownReason2"]);
                            wrksht.Cells[row, 20] = Convert.ToString(rs["Operator"]);
                            interior = interior + 1;
                        }

                        wrksht.get_Range(wrksht.Cells[row1, 2], wrksht.Cells[row, 2]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 3], wrksht.Cells[row, 3]).Merge(false);

                        wrksht.Cells[row, 4] = Convert.ToString(rs["Component"]);
                        wrksht.Cells[row, 5] = Convert.ToString(rs["RunTime"]);
                        wrksht.Cells[row, 6] = Convert.ToString(rs["Target"]);
                        wrksht.Cells[row, 7] = Convert.ToString(rs["CountShift1"]);
                        wrksht.Cells[row, 8] = Convert.ToString(rs["CountShift2"]);
                        wrksht.Cells[row, 9] = Convert.ToString(rs["CountShift3"]);

                        wrksht.Cells[row, 10] = Convert.ToString(Convert.ToUInt32(rs["CountShift1"]) + Convert.ToUInt32(rs["CountShift2"]) + Convert.ToUInt32(rs["CountShift3"]));
                        wrksht.Cells[row, 11] = Convert.ToString(rs["frmtStdCycletime"]); // g: col + 1, 10th col is total of 3 shifts
                        wrksht.Cells[row, 12] = Convert.ToString(rs["frmtActualCycletime"]);
                        wrksht.Cells[row, 13] = Convert.ToString(rs["Cyclefficiency"]);
                        wrksht.get_Range(wrksht.Cells[row1, 13], wrksht.Cells[row, 13]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 14], wrksht.Cells[row, 14]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 15], wrksht.Cells[row, 15]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 16], wrksht.Cells[row, 16]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 17], wrksht.Cells[row, 17]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 18], wrksht.Cells[row, 18]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 19], wrksht.Cells[row, 19]).Merge(false);
                        wrksht.get_Range(wrksht.Cells[row1, 20], wrksht.Cells[row, 20]).Merge(false);
                        if (interior % 2 == 0)
                        {
                            wrksht.get_Range(wrksht.Cells[row, 1], wrksht.Cells[row1, 20]).Interior.ColorIndex = 40;
                        }
                        row++;
                    }
                }
                else
                {
                    Logger.WriteDebugLog("No records found.");
                }

                if (rs != null)
                {
                    rs.Close();
                }
                wrksht.Columns.AutoFit();

                //wrkbk.Close(true, misValue, misValue);
                wrkbk.SaveAs(dst, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange,
                        false, false, Missing.Value, Missing.Value, Missing.Value);
                if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");
                Thread.Sleep(2000);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                if (xlApp != null)
                {
                    releaseObject(wrkbk);
                    releaseObject(wrksht);
                    releaseObject(xlApp);
                }
                if (pid != 0) KillSpecificExcelFileProcess(pid);
            }
            return true;
        }


        public static bool ExportMOReport(string strReportFile, string ExportPath, string ExportedReportFile,
           int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime,
           string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string CompanyName, bool MachineAE)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wrkbk = null;
            Excel.Worksheet wrksht = null;
            object misValue = System.Reflection.Missing.Value;
            SqlDataReader rs = null;
            //sttime = "2018-Jan-20 06:00:00 AM"; //g: testdates
            //ndtime = "2018-Jan-21 06:00:00 AM";
            int pid = 0;

            try
            {
                string src, dst = string.Empty;//Globally Used  
                string plantname = plantid;
                string SDate = string.Format("{0:yyyy-MMM-dd hh:mm:ss tt}", sttime);
                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\SM_GetMODetails.xls";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = ExportPath + @"MODetailsRportWeekly_" + plantname + "_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(sttime)) + ".xls";//string.Format("{0:hh-mm-ss MMM-yyyy}", DT) + ".xls";


                Logger.WriteDebugLog("Generating reports for " + sttime + " : " + ndtime);
                rs = AccessReportData.GetMoReport(DateTime.Parse(sttime), DateTime.Parse(ndtime), plantid, MachineId, "");

                try
                {
                    File.Copy(src, dst, true);
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }

                if (!File.Exists(dst))
                {
                    if (rs != null) rs.Close();
                    return false;
                }
                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;
                int a = GetWindowThreadProcessId(xlApp.Hwnd, out pid);
                wrkbk = xlApp.Workbooks.Open(dst, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);

                int interior = 0;
                string Machineref = "firstrecord", DownReasonRef = "firstrecord";
                int row = 4, col = 1;
                int RowStart = row;
                int RowMCStart = row;
                wrksht.Cells[2, 2] = sttime;
                wrksht.Cells[2, 7] = ndtime; // g:

                while (rs.Read())
                {
                    wrksht.Cells[row, col] = rs["Sdate"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["Pdate"].ToString(); // g:
                    col = col + 1; // g:
                    wrksht.Cells[row, col] = rs["CellNo"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["MCNo"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["MONo"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["ItemCode"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["EmployeeName"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["ActualCount"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["MOQuantity"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["MOSettingTime"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["MORunningTime"].ToString();
                    col = col + 1;
                    wrksht.Cells[row, col] = rs["TotalCycletime"].ToString();
                    col = col + 2;
                    wrksht.Cells[row, col] = rs["ActualTime"].ToString();
                    col = col + 4;

                    if (DownReasonRef != rs["Remarks1"].ToString() && Machineref == rs["MCNo"].ToString())
                    {
                        wrksht.Cells[row, col] = rs["Remarks1"].ToString();
                        DownReasonRef = rs["Remarks1"].ToString();
                        RowStart = row;
                    }

                    if (DownReasonRef != rs["Remarks1"].ToString() || Machineref != rs["MCNo"].ToString())
                    {
                        wrksht.Cells[row, col] = rs["Remarks1"].ToString();
                        DownReasonRef = rs["Remarks1"].ToString();
                        RowStart = row;
                    }

                    if (DownReasonRef == rs["Remarks1"].ToString() && Machineref == rs["MCNo"].ToString())
                    {
                        wrksht.get_Range(wrksht.Cells[RowStart, 17], wrksht.Cells[row, 17]).MergeCells = true;
                        wrksht.get_Range(wrksht.Cells[RowStart, 18], wrksht.Cells[row, 18]).MergeCells = true;
                    }

                    if (Machineref != rs["MCNo"].ToString())
                    {
                        RowMCStart = row;
                        interior = interior + 1;
                    }
                    if (interior % 2 == 0)
                    {
                        wrksht.get_Range(wrksht.Cells[RowMCStart, 1], wrksht.Cells[row, 18]).Interior.ColorIndex = 40;
                    }

                    Machineref = rs["MCNo"].ToString();
                    col = 1;
                    row = row + 1;

                }
                if (rs != null)
                {
                    rs.Close();
                }

                Excel.Worksheet xlWorkSheetFocus = (Excel.Worksheet)wrkbk.Worksheets.get_Item(1);
                xlWorkSheetFocus.get_Range("A4", "S" + (row - 1)).Borders.ColorIndex = 0; // g:
                xlWorkSheetFocus.get_Range("A4", "AS" + row).EntireColumn.AutoFit();      // g:
                Range tt = xlWorkSheetFocus.get_Range("A1", "A1");
                tt.Select();
                xlWorkSheetFocus.Activate();
                if (wrkbk != null) wrkbk.Close(true, misValue, misValue);
                if (xlApp != null) xlApp.Quit();
                Logger.WriteDebugLog("Report generated sucessfully.");
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Report generation Failed Error: " + ex.ToString());
                return false;
            }
            finally
            {
                if (rs != null && rs.IsClosed == false) rs.Close();
                if (xlApp != null)
                {
                    releaseObject(wrksht);
                    releaseObject(wrkbk);
                    releaseObject(xlApp);
                    if (pid != 0) KillSpecificExcelFileProcess(pid);
                }
            }
            return true;
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        // g:
        internal static void ExportPMReportShantiIron(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
                  string MachineId, string operators, string sttime,
                  string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                  string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath, @"SM_PM_Report_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx");
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillPMReportExcelSheetData(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }

        private static void FillPMReportExcelSheetData(string strtTime, string endTimez, string dst, string src, string machineId,
            string plantId, bool Email_Flag, string Email_List_To,
            string Email_List_CC, string Email_List_BCC, string exportFileName)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();

            try
            {
                FileInfo newFile = new FileInfo(dst);
                FileInfo tempFile = new FileInfo(src);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                Dictionary<string, List<string>> catAndSubCat = new Dictionary<string, List<string>>();
                Dictionary<string, int> catrow = new Dictionary<string, int>();
                Dictionary<string, int> subcatrow = new Dictionary<string, int>();
                Dictionary<string, int> monthcol = new Dictionary<string, int>{
                    {"1", 1},
                    {"2", 2},
                    {"3", 3},
                    {"4", 4},
                    {"5", 5},
                    {"6", 6},
                    {"7", 7},
                    {"8", 8},
                    {"9", 9},
                    {"10", 10},
                    {"11", 11},
                    {"12", 12}
                };

                int month = Convert.ToDateTime(strtTime).Month;
                foreach (var k in monthcol.Keys.ToList())
                {
                    monthcol[k] = ((monthcol[k] + 12 - month) % 12 + 5); // with added offset to fill in the sheet
                }

                string strMacName = string.Empty;
                int row = 6;
                int col = 5;
                int cnt = 1;
                string startTime = strtTime, endTime = endTimez;

                System.Data.DataTable dt = AccessReportData.GetPMReport(startTime, endTime, machineId);

                List<string> machineLst = new List<string>();
                int lastrow = 0;


                foreach (DataRow rdr in dt.Rows)
                {
                    try
                    {
                        if (strMacName != rdr["machineid"].ToString())
                        {
                            strMacName = rdr["machineid"].ToString();
                            machineLst.Add(strMacName);
                            ws = excelPackage.Workbook.Worksheets.Add(strMacName, excelPackage.Workbook.Worksheets[1]);
                            catAndSubCat = AccessReportData.GetCatAndSubCat(strMacName);
                            row = 6;
                            col = 5;
                            FillTempl(ref ws, ref row, ref col, ref catrow, ref subcatrow, ref catAndSubCat, ref strtTime, ref lastrow);

                            ws.Cells["D3"].Value = strMacName;
                            ws.Cells["D4"].Value = rdr["Machine"];
                            ws.Cells["P3"].Value = DateTime.Now.ToString("dd-MM-yy");
                            ws.Cells["P4"].Value = DateTime.Now.ToString("HH:mm");
                            ws.Name = rdr["machineid"].ToString();
                            ws.Cells[subcatrow[rdr["SubCategory"].ToString()], monthcol[rdr["mon"].ToString()]].Value =
                                rdr["Record"].ToString().Equals("NOT OK", StringComparison.OrdinalIgnoreCase) ?
                                rdr["Reason"].ToString() :
                                rdr["Record"].ToString();
                            ws.Cells[6, monthcol[rdr["mon"].ToString()]].Value = Convert.ToDateTime(rdr["Starttime"].ToString()).ToString("dd-MMM-yy");
                            ws.Cells[6, monthcol[rdr["mon"].ToString()]].Style.Font.Bold = true;
                            cnt += 1;
                        }
                        else
                        {
                            ws.Cells[subcatrow[rdr["SubCategory"].ToString()], monthcol[rdr["mon"].ToString()]].Value =
                                rdr["Record"].ToString().Equals("NOT OK", StringComparison.OrdinalIgnoreCase) ?
                                rdr["Reason"].ToString() :
                                rdr["Record"].ToString();
                            ws.Cells[6, monthcol[rdr["mon"].ToString()]].Value = Convert.ToDateTime(rdr["Starttime"].ToString()).ToString("dd-MMM-yy");
                            ws.Cells[6, monthcol[rdr["mon"].ToString()]].Style.Font.Bold = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteDebugLog(string.Format("Possibly keys not found: Subcategory: {0}, month: {1}", rdr["SubCategory"].ToString(), rdr["mon"].ToString()));
                    }
                }


                List<string> reasonLst = new List<string>();
                // modify cells to replace with "not ok" and to move reason to last column
                foreach (string mach in machineLst)
                {
                    ws = excelPackage.Workbook.Worksheets[mach];
                    for (row = 7; row < 22; row++)
                    {
                        for (col = 5; col < 17; col++)
                        {
                            if (ws.Cells[row, col].Value != null)
                            {
                                if (!(ws.Cells[row, col].Value.ToString().Equals("OK") || ws.Cells[row, col].Value.ToString().Equals("")))
                                {
                                    reasonLst.Add(string.Format("({0}) {1}", ws.Cells[6, col].Value.ToString(), ws.Cells[row, col].Value.ToString()));
                                    ws.Cells[row, col].Value = "NOT OK";
                                    ws.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                }
                                if (ws.Cells[row, col].Value.ToString().Equals("OK"))
                                {
                                    ws.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                }
                            }
                        }
                        ws.Cells[row, 17].Value = string.Join(Environment.NewLine, reasonLst.ToArray());
                        ws.Cells[row, 17].Style.WrapText = true;

                        reasonLst.Clear();
                    }
                }

                if (machineLst.Count > 0)
                    excelPackage.Workbook.Worksheets.Delete(excelPackage.Workbook.Worksheets[1]);
                SetPrinterSettings(ws);

                excelPackage.SaveAs(newFile);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error creating Excel File." + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }


        internal static void ExportDailyProductionReportDayWiseShantiIron(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
                  string MachineId, string operators, string sttime,
                  string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                  string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_DailyOEEByShift_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExcelSheetData(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }
        //Vasavi Added
        internal static void ExportProductionRExportCockpitProductionReportShantiIron(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
                  string MachineId, string operators, string sttime,
                  string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                  string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_CockpitDataReport_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExcelSheet(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }


        //Vasavi Added
        internal static void ExportProductionCountBySlNo(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
                string MachineId, string operators, string sttime,
                string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_ShiftWiseProductionDetails_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExcelSheetForProductionCountBySlno(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }



        internal static void ExportMachinewiseAlarmReport(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
                string MachineId, string operators, string sttime,
                string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
                string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_MachineAlarmReport_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExportMachinewiseAlarmReport(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }


        //Vasavi Added
        internal static void ExportMonthwiseOEEReport(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
           string MachineId, string operators, string sttime,
           string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_MonthlyOEEReport_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExportMonthwiseOEEReport(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }


        //Vasavi Added
        internal static void ExportMachineDownTimeMatrix(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
           string MachineId, string operators, string sttime,
           string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC)
        {
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_MachineDownTimeMatrix_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillExportMachineDownTimeMatrixReport(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }

        internal static void ExportMachineDownTimeMatrix_Advik(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
           string MachineId, string operators, string sttime,
           string PlantId,string cellID, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC)
        {
            string dst = string.Empty; bool isDataAvailable = false;
            int LimitData = 11, screenWidth=1400, screenHeight=800;
            strReportFile = _appPath + "\\Reports\\MachineDownTimeReport_Advik.xlsx";
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath ,string.Format("SM_MachineDownTimeMatrix_{0:ddMMMyyyy_HHmmss}.xlsx", DateTime.Parse(strtTime)));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ExcelWorksheet ws2 = excelPackage.Workbook.Worksheets[2];

                //foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
                //{
                // sheet.Name = sheet.Name.Replace("5", LimitData);
                //}
                ws.Cells["C2"].Value = strtTime;
                ws.Cells["E2"].Value = endTime;
                ws.Cells["G2"].Value = PlantId == "" ? "All" : PlantId;
                int r = 7, count = 0, c = 2;
                List<string> lstMachineNames = new List<string>();
                List<string> lstMacTotal = new List<string>();
                List<string> lstMacFreqTotal = new List<string>();
                string mxkSectohhmmss = string.Empty;

                string PrevDownFreq = string.Empty;
                string DownIDTotal = string.Empty;
                string downtime = string.Empty;
                string PrevMachine = string.Empty;
                int range = LimitData + 6;
                string Prevdown = string.Empty; string Machinename = "";

                if (ConfigurationManager.AppSettings["sonapages"].Equals("1"))
                {
                    Machinename = "machineDescription";
                }
                else
                    Machinename = "MachineID";

                System.Data.DataTable dt = AccessReportData.MachineDownTimeMatrix(Convert.ToDateTime(strtTime),Convert.ToDateTime(endTime), MachineId, PlantId, "", 0, "DTime", cellID, "s_GetDownTimeMatrixfromAutoData", "");
                if (dt != null && dt.Rows.Count > 0)
                {
                    isDataAvailable = true;
                    #region "Header Part Machine "
                    foreach (DataRow rdr in dt.Rows)
                    {
                        if (count == 0)
                        {
                            Prevdown = rdr["DownCode"].ToString();
                            PrevDownFreq = rdr["DownCode"].ToString();
                            PrevMachine = rdr[Machinename].ToString();
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            if (ConfigurationManager.AppSettings["sonapages"].Equals("1", StringComparison.OrdinalIgnoreCase))
                            {
                                ws.Cells[5, c].Value = rdr["MachineID"].ToString();
                                ws.Cells[5, c, 5, c + 1].Merge = true;
                            }
                            ws.Cells[6, c].Value = PrevMachine;
                            ws.Cells[6, c, 6, c + 1].Merge = true;
                            ws.Cells[7, c].Value = "Down Time";
                            ws.Cells[7, c + 1].Value = "Frequency";
                            lstMacTotal.Add(rdr["TotalMachine"].ToString());
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        if (PrevMachine != rdr[Machinename].ToString())
                        {
                            c = c + 2;
                            if (ConfigurationManager.AppSettings["sonapages"].Equals("1", StringComparison.OrdinalIgnoreCase))
                            {
                                ws.Cells[5, c].Value = rdr["MachineID"].ToString();
                                ws.Cells[5, c, 5, c + 1].Merge = true;
                            }
                            ws.Cells[6, c].Value = rdr[Machinename].ToString();
                            ws.Cells[6, c, 6, c + 1].Merge = true;
                            ws.Cells[7, c].Value = "Down Time";
                            ws.Cells[7, c + 1].Value = "Frequency";
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            lstMacTotal.Add(rdr["TotalMachine"].ToString());
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        else if (PrevMachine == rdr[Machinename].ToString() && count != 0)
                        {
                            break;
                        }
                        count++;
                    }
                    #endregion
                    #region "Machine Value Define Frist Sheet"
                    r = 8;
                    c = 2;
                    Prevdown = "";
                    TimeSpan timeSpan = TimeSpan.MinValue;
                    foreach (DataRow rdr in dt.Rows)
                    {
                        if (Prevdown == "" || Prevdown == rdr["DownCode"].ToString())
                        {
                            //ws.Cells[r, 1].Value = rdr["DownCode"].ToString();
                            ws.Cells[r, c].Value = Convert.ToDecimal(rdr["DownTime"].ToString()) / 86400;
                            ws.Column(c + 1).Style.Numberformat.Format = "0";
                            ws.Cells[r, c + 1].Value = Convert.ToDecimal(rdr["DownFreq"].ToString());
                            //ws.Cells[r, c + 1].Style.Numberformat.Format = "Number";
                        }
                        else if (Prevdown != rdr["DownCode"].ToString())
                        {
                            r = r + 1;
                            c = 2;
                            ws.Cells[r, c].Value = Convert.ToDecimal(rdr["DownTime"].ToString()) / 86400;
                            ws.Column(c + 1).Style.Numberformat.Format = "0";
                            ws.Cells[r, c + 1].Value = Convert.ToDecimal(rdr["DownFreq"].ToString());
                        }
                        c = c + 2;
                        ws.Cells[r, 1].Value = rdr["DownCode"].ToString();
                        Prevdown = rdr["DownCode"].ToString();
                    }
                    c = 2;
                    ws.Cells[r + 2, 1].Value = "Total";
                    foreach (string value in lstMacTotal)
                    {
                        if (Convert.ToInt32(value) > 0)
                        {
                            timeSpan = TimeSpan.FromSeconds(Convert.ToDouble(value));
                            string answer = string.Format("{0:00}:{1:00}:{2:00}",
                            (int)timeSpan.TotalHours,
                            timeSpan.Minutes,
                            timeSpan.Seconds);
                            //ws.Cells[r + 2, c].Value = Convert.ToDecimal(value) / 86400;
                            ws.Cells[r + 2, c].Style.Numberformat.Format = "General";
                            ws.Cells[r + 2, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[r + 2, c].Value = answer;
                        }
                        c = c + 2;

                    }
                    c = 3;
                    foreach (string values in lstMacFreqTotal)
                    {
                        //-------Start --------------
                        int total = 0;
                        if (int.TryParse(values, out total))
                            ws.Cells[r + 2, c].Value = total;
                        //------END---------------------
                        c = c + 2;
                    }
                    #endregion
                    #region Charts

                    range = LimitData + 7;
                    var chart11 = ws2.Drawings["Chart 1"] as ExcelBarChart;
                    chart11.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;
                    chart11.YAxis.Format = "[h]:mm:ss;@";
                    chart11.SetSize(screenWidth, screenHeight);
                    chart11.SetPosition(10, 22);

                    chart11.Title.Text = "Down Time Comparison Graph";
                    chart11.YAxis.Title.Text = "Down Time";
                    for (int i = 2; i < c - 1; i = i + 2)
                    {
                        ExcelChartSerie aa = null;
                        if (r>range)
                        {
                             aa = chart11.Series.Add(ws.Cells[8, i, range, i], ws.Cells[8, 1, range, 1]);
                        }
                        else
                        {
                             aa = chart11.Series.Add(ws.Cells[8, i, r, i], ws.Cells[8, 1, r, 1]);
                        }
                        aa.HeaderAddress = new ExcelAddress("'Time-wise'!" + GetExcelColumnName(i) + "6");
                    }
                    var barchart = ws2.Drawings["Chart 2"] as ExcelBarChart;
                    barchart.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;
                    // var series = barchart.Series[0];
                    barchart.Title.Text = "Frequency Comparison Graph";
                    barchart.YAxis.Title.Text = "Frequency";
                    barchart.YAxis.Format = "00";
                    barchart.SetSize(screenWidth, screenHeight);
                    barchart.SetPosition(850, 22);

                    for (int i = 3; i < c+1 ; i = i + 2)
                    {
                        ExcelChartSerie aa = null;
                        if (r>range)
                        {
                             aa = barchart.Series.Add(ws.Cells[8, i, range, i], ws.Cells[8, 1, range, 1]);
                        }
                        else
                        {
                             aa = barchart.Series.Add(ws.Cells[8, i, r, i], ws.Cells[8, 1, r, 1]);
                        }
                        aa.HeaderAddress = new ExcelAddress("'Time-wise'!" + GetExcelColumnName(i-1 ) + "6");
                    }
                    Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
                    #endregion
                    excelPackage.SaveAs(newFile);
                    if (isDataAvailable)
                    {
                        SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                        Logger.WriteDebugLog("File Mailed Successfully");
                    }
                    
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }

        internal static void ExportDailyProductionandRejectionReport(string strtTime, string endTime, string Shift, string strReportFile, string ExportPath, string ExportedReportFile,
           string MachineId, string operators, string sttime,
           string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, int ShiftID)
        {
            string _appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                string reppath = Path.Combine(_appPath, "Reports");

                dst = Path.Combine(reppath, string.Format("Temp\\SM_DailyProductionandRejectionReport_{0}.xlsx", string.Format("{0:MMMyyyy_}", DateTime.Parse(AccessReportData.GetLogicalDayStart(strtTime)))));

                if (!File.Exists(dst))
                {
                    if (!Directory.Exists(Path.Combine(_appPath, "Reports","Temp")))
                    {
                        Directory.CreateDirectory(Path.Combine(_appPath, "Reports", "Temp"));
                    }
                    File.Copy(strReportFile, dst, true);
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                }
                //Logger.WriteDebugLog("Before Writing into excel ");

                FillExportProductionAndRejectionReport(strtTime, endTime, Shift, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile, ShiftID);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");

                // Logger.WriteDebugLog("After writing into temp folder ");
                string genertadReport = Path.Combine(ExportPath, string.Format("SM_DailyProductionandRejectionReport_{0}.xlsx", string.Format("{0:MMMyyyy_}", DateTime.Parse(AccessReportData.GetLogicalDayStart(strtTime)))));

                if (!File.Exists(genertadReport))
                {
                    File.Copy(dst, genertadReport, true);
                }
                Logger.WriteDebugLog("After writing into reports folder ");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }

        }

        private static void FillExportMachineDownTimeMatrixReport(string strtTime, string endTime, string dst, string strReportFile, string MachineId, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string ExportedReportFile)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }
                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }

                string CmbDowntimeformat = "ss";
                string strPlantName = string.Empty;
                string tempmachineid = string.Empty;

                List<string> lstMachineNames = new List<string>();
                List<string> lstMacTotal = new List<string>();
                List<string> lstMacFreqTotal = new List<string>();
                string mxkSectohhmmss = string.Empty;
                int r = 6, c = 1;
                string startTime = strtTime;


                string DownID = string.Empty;
                string PrevDownFreq = string.Empty;
                string DownIDTotal = string.Empty;
                string downtime = string.Empty;

                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetDownTimeMatrixfromAutoData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@DownID", "");
                cmd.Parameters.AddWithValue("@OperatorID", "");
                cmd.Parameters.AddWithValue("@ComponentID", "");
                cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@MatrixType", "DTime");
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Excludedown", 0);
                int count = 0;
                string PrevMachine = string.Empty;
                strPlantName = string.Empty;
                string Prevdown = string.Empty;
                SqlDataReader rdr = cmd.ExecuteReader();

                r = 7;
                c = 2;
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ExcelWorksheet ws2 = excelPackage.Workbook.Worksheets[2];
                ExcelWorksheet ws3 = excelPackage.Workbook.Worksheets[3];
                ExcelWorksheet ws4 = excelPackage.Workbook.Worksheets[4];
                ExcelWorksheet ws5 = excelPackage.Workbook.Worksheets[5];
                ExcelWorksheet ws6 = excelPackage.Workbook.Worksheets[6];
                ExcelWorksheet ws7 = excelPackage.Workbook.Worksheets[7];

                #region Sheet1
                #region
                if (rdr.HasRows)
                {
                    //  ws.Name = "Time-wise";
                    ws.Cells[4, 5].Value = "";
                    ws.Cells[2, 6].Value = strtTime;
                    ws.Cells[2, 10].Value = endTime;

                    while (rdr.Read())
                    {
                        if (count == 0)
                        {
                            Prevdown = rdr["DownCode"].ToString();
                            PrevDownFreq = rdr["DownCode"].ToString();
                            PrevMachine = rdr["MachineID"].ToString();
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            ws.Cells[6, c].Value = PrevMachine;
                            ws2.Cells[6, c].Value = PrevMachine;
                            ws3.Cells[6, c].Value = PrevMachine;
                            lstMacTotal.Add(rdr["TotalMachine"].ToString());
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        if (PrevMachine != rdr["MachineID"].ToString())
                        {
                            c = c + 1;
                            ws.Cells[6, c].Value = rdr["MachineID"].ToString();
                            ws2.Cells[6, c].Value = rdr["MachineID"].ToString();
                            ws3.Cells[6, c].Value = rdr["MachineID"].ToString();
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            lstMacTotal.Add(rdr["TotalMachine"].ToString());
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        else if (PrevMachine == rdr["MachineID"].ToString() && count != 0)
                        {
                            break;
                        }
                        count++;
                    }

                }
                rdr.Close();
                #endregion

                #region firstSheetData
                cmd = new SqlCommand(@"[s_GetDownTimeMatrixfromAutoData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@DownID", "");
                cmd.Parameters.AddWithValue("@OperatorID", "");
                cmd.Parameters.AddWithValue("@ComponentID", "");
                cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@MatrixType", "DTime");
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Excludedown", 0);

                r = 7;
                c = 2;
                Prevdown = "";
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {

                        if (Prevdown == "" || Prevdown == rdr["DownCode"].ToString())
                        {
                            //   ws.Cells[r, 1].Value = rdr["DownCode"].ToString();
                            ws.Cells[r, c].Value = Convert.ToDecimal(rdr["DownTime"].ToString()) / 86400;


                        }
                        else if (Prevdown != rdr["DownCode"].ToString())
                        {

                            r = r + 1;
                            c = 2;
                            ws.Cells[r, c].Value = Convert.ToDecimal(rdr["DownTime"].ToString()) / 86400;
                        }
                        c = c + 1;
                        ws.Cells[r, 1].Value = rdr["DownCode"].ToString();
                        Prevdown = rdr["DownCode"].ToString();

                    }
                }
                rdr.Close();
                c = 2;
                ws.Cells[r + 2, 1].Value = "Total";

                foreach (string value in lstMacTotal)
                {
                    if (Convert.ToInt32(value) > 0)
                    {
                        ws.Cells[r + 2, c].Value = Convert.ToDecimal(value) / 86400;
                    }
                    c = c + 1;

                }
                #endregion
                #endregion

                ws2.Cells[2, 6].Value = strtTime;
                ws2.Cells[2, 10].Value = endTime;
                #region Sheet2
                cmd = new SqlCommand(@"[s_GetDownTimeMatrixfromAutoData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@DownID", "");
                cmd.Parameters.AddWithValue("@OperatorID", "");
                cmd.Parameters.AddWithValue("@ComponentID", "");
                cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@MatrixType", "DTime");
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Excludedown", 0);

                r = 7;
                c = 2;
                PrevDownFreq = "";
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        if (PrevDownFreq == "" || PrevDownFreq == rdr["DownCode"].ToString())
                        {

                            ws2.Cells[r, c].Value = Convert.ToInt32(rdr["DownFreq"].ToString());

                        }
                        else
                        {
                            r = r + 1;
                            c = 2;
                            ws2.Cells[r, c].Value = Convert.ToInt32(rdr["DownFreq"].ToString());
                        }
                        c = c + 1;
                        ws2.Cells[r, 1].Value = rdr["DownCode"].ToString();
                        PrevDownFreq = rdr["DownCode"].ToString();
                    }
                }
                rdr.Close();
                c = 2;
                ws2.Cells[r + 2, 1].Value = "Total";

                foreach (string value in lstMacFreqTotal)
                {
                    ws2.Cells[r + 2, c].Value = value;
                    c = c + 1;

                }
                #endregion

                var chart11 = (ExcelBarChart)ws3.Drawings.AddChart("DownTime Comparison Graph", eChartType.ColumnClustered);
                chart11.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;

                chart11.SetSize(1000, 500);
                chart11.SetPosition(10, 22);
                chart11.Title.Text = "Down Time Comparison Graph";
                for (int i = 2; i < c; i++)
                {
                    ExcelChartSerie aa = chart11.Series.Add(ws.Cells[7, i, 11, i], ws.Cells[7, 1, 11, 1]);

                    aa.HeaderAddress = new ExcelAddress("'MCs by Top-5 Downs'!" + GetExcelColumnName(i) + "6");
                }

                var barchart = ws4.Drawings["Chart 2"] as ExcelBarChart;
                var series = barchart.Series[0];
                barchart.Title.Text = "Down Time Comparison Graph";
                barchart.SetSize(1000, 500);
                barchart.SetPosition(10, 22);
                series.XSeries = ws.Cells[6, 2, 6, c].FullAddress;
                series.Series = ws.Cells[7, 2, 7, c].FullAddress;


                ws5.Cells[2, 6].Value = strtTime;
                ws5.Cells[2, 10].Value = endTime;

                #region Sheet5
                cmd = new SqlCommand(@"[s_GetDownTimeMatrixfromAutoData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@DownID", "");
                cmd.Parameters.AddWithValue("@OperatorID", "");
                cmd.Parameters.AddWithValue("@ComponentID", "");
                cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@MatrixType", "DFreq");
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Excludedown", 0);

                rdr = cmd.ExecuteReader();

                r = 7;
                c = 2;
                count = 0;

                #region
                if (rdr.HasRows)
                {

                    ws5.Cells[4, 5].Value = "";
                    ws5.Cells[2, 7].Value = strtTime;
                    ws5.Cells[2, 10].Value = endTime;
                    lstMacFreqTotal.Clear();
                    lstMachineNames.Clear();
                    lstMacTotal.Clear();
                    while (rdr.Read())
                    {
                        if (count == 0)
                        {
                            PrevDownFreq = rdr["DownCode"].ToString();
                            PrevMachine = rdr["MachineID"].ToString();
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            ws5.Cells[6, c].Value = PrevMachine;
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        if (PrevMachine != rdr["MachineID"].ToString())
                        {
                            c = c + 1;
                            ws5.Cells[6, c].Value = rdr["MachineID"].ToString();
                            lstMachineNames.Add(rdr["MachineID"].ToString());
                            lstMacFreqTotal.Add(rdr["TotalMachineFreq"].ToString());
                        }
                        else if (PrevMachine == rdr["MachineID"].ToString() && count != 0)
                        {
                            break;
                        }
                        count++;
                    }

                }
                rdr.Close();
                #endregion
                cmd = new SqlCommand(@"[s_GetDownTimeMatrixfromAutoData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@DownID", "");
                cmd.Parameters.AddWithValue("@OperatorID", "");
                cmd.Parameters.AddWithValue("@ComponentID", "");
                cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
                cmd.Parameters.AddWithValue("@MatrixType", "DFreq");
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Excludedown", 0);
                r = 7;
                c = 2;
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        if (PrevDownFreq == "" || PrevDownFreq == rdr["DownCode"].ToString())
                        {

                            ws5.Cells[r, c].Value = Convert.ToInt32(rdr["DownFreq"].ToString());
                        }
                        else
                        {
                            r = r + 1;
                            c = 2;
                            ws5.Cells[r, c].Value = Convert.ToInt32(rdr["DownFreq"].ToString());

                        }
                        c = c + 1;
                        ws5.Cells[r, 1].Value = rdr["DownCode"].ToString();
                        PrevDownFreq = rdr["DownCode"].ToString();
                    }
                }
                rdr.Close();
                c = 2;
                ws5.Cells[r + 2, 1].Value = "Total";

                foreach (string value in lstMacFreqTotal)
                {
                    ws5.Cells[r + 2, c].Value = value;
                    c = c + 1;

                }
                #endregion

                var chartfreqWise = (ExcelBarChart)ws6.Drawings.AddChart("Down Frequency Comparison Graph", eChartType.ColumnClustered);
                chartfreqWise.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;


                chartfreqWise.Title.Text = "Down Frequency Comparison Graph";
                chartfreqWise.YAxis.Title.Text = "Down Frequency";

                chartfreqWise.SetSize(1000, 500);
                chartfreqWise.SetPosition(10, 22);
                for (int i = 2; i < c; i++)
                {
                    ExcelChartSerie aa = chartfreqWise.Series.Add(ws5.Cells[7, i, 11, i], ws5.Cells[7, 1, 12, 1]);
                    aa.HeaderAddress = new ExcelAddress("'Freq-wise'!" + GetExcelColumnName(i) + "6");
                }

                var BarChartForFreq = ws7.Drawings["Chart 1"] as ExcelBarChart;
                var series1 = BarChartForFreq.Series[0];
                BarChartForFreq.Title.Text = "Down Frequency Comparison Graph";
                BarChartForFreq.YAxis.Title.Text = "Down Frequency";
                BarChartForFreq.SetSize(1000, 500);
                BarChartForFreq.SetPosition(10, 22);
                series1.XSeries = ws5.Cells[6, 2, 6, c].FullAddress;
                series1.Series = ws.Cells[7, 2, 7, c].FullAddress;

                //ws.Name = "Time-wise";
                //ws2.Name = "Time-wise Freq";
                SetPrinterSettings(ws);
                SetPrinterSettings(ws2);
                excelPackage.SaveAs(newFile);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        internal static int GetMchRateForMachine(string machineid)
        {
            List<string> list = new List<string>();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            int mchrate = 0;
            SqlDataReader rdr = null;
            string sqlQuery = string.Empty;
            try
            {
                sqlQuery = @"select mchrrate from machineinformation where machineid=@machineid";
                cmd = new SqlCommand(sqlQuery, conn);
                cmd.Parameters.AddWithValue("@machineid", machineid);
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandTimeout = 120;
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (!Convert.IsDBNull(rdr["mchrrate"]))
                    {
                        mchrate = Convert.ToInt32(rdr["mchrrate"]);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
                throw;
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return mchrate;
        }



        #region extraz
        //Vasavi
        internal static void MoneyVals(double DwnTime, double mchval, out double MoneyVal)
        {
            long tmsec = 0;
            double tmhour = 0;
            MoneyVal = 0.0;

            tmsec = Convert.ToInt64(DwnTime);
            if (tmsec != 0)
            {
                tmhour = tmsec / 3600;
            }
            else
            {
                tmhour = 0;
            }

            MoneyVal = tmhour * mchval;
            MoneyVal = Math.Round(MoneyVal, 2);
        }
        internal static void GetmxkSectohhmmss(double sec, out string mxkSectohhmmss)
        {
            double hours = 0.0; long hour = 0; double minutes = 0.0; int minute = 0;
            int second = 0;
            string strhour = string.Empty;
            string strminute = string.Empty;
            string strsecond = string.Empty;
            mxkSectohhmmss = string.Empty;
            try
            {
                if (sec == 0.0)
                {
                    mxkSectohhmmss = "00:00:00";
                }
                else
                {
                    hours = Math.Abs(Math.Round(sec / 3600, 5));
                    hour = Convert.ToInt32(hours);
                    minutes = Math.Round(((hours - hour) * 60), 3);
                    minute = Convert.ToInt32(minutes);
                    second = Convert.ToInt32((minutes - minute) * 60);
                    if (sec < 0)
                    {
                        hour = hour * -1;
                    }
                    if (hour < 0)
                    {
                        strhour = "-" + hour;
                    }
                    else if (hour <= 9)
                    {
                        strhour = "0" + hour;
                    }
                    else
                    {
                        strhour = hour.ToString();
                    }

                    if (minute <= 9)
                    {
                        strminute = "0" + minute;
                    }
                    else
                    {
                        strminute = minute.ToString();
                    }
                    if (second <= 9)
                    {
                        strsecond = "0" + second;
                    }
                    else
                    {
                        strsecond = second.ToString();
                    }
                    mxkSectohhmmss = strhour + ":" + strminute + ":" + strsecond;

                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Getting mxkSectohhmmss value..!!\n " + ex.Message);
                throw;
            }
        }
        #endregion

        private static void FillExportMonthwiseOEEReport(string strtTime, string endTime, string dst, string strReportFile, string MachineId, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string ExportedReportFile)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                string strPlantName = string.Empty;
                int row = 5, col = 1;
                string startTime = strtTime;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetMonthwiseEfficiencyFromAutodata_Shanthi]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", startTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@Param", "");
                strPlantName = string.Empty;
                SqlDataReader rdr = cmd.ExecuteReader();

                DateTime ST = Convert.ToDateTime(startTime);
                DateTime PreMonth = Convert.ToDateTime(startTime);
                PreMonth = PreMonth.AddMonths(-1);
                string sMonth = ST.ToString("MMM");

                string PreviousMonth = PreMonth.ToString("MMM");
                double CurrentMonthAvg = 0.0;

                double PreMonthAvg = 0.0;

                if (rdr.HasRows)
                {
                    FileInfo newFile = new FileInfo(dst);
                    ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                    ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                    ws.Cells[5, 23].Value = sMonth;
                    ws.Cells[6, 23].Value = PreviousMonth;


                    while (rdr.Read())
                    {
                        ws.Cells[row, 1].Value = rdr["MachineID"];
                        ws.Cells[row, 1].Style.Font.Bold = true;
                        ws.Cells[row, 2].Value = rdr["Components"];
                        ws.Cells[row, 3].Value = rdr["RejCount"];
                        ws.Cells[row, 4].Value = rdr["ProductionEfficiency"];
                        ws.Cells[row, 5].Value = rdr["AvailabilityEfficiency"];
                        ws.Cells[row, 6].Value = rdr["QualityEfficiency"];
                        ws.Cells[row, 7].Value = rdr["OverAllEfficiency"];
                        ws.Cells[row, 8].Value = rdr["AvgOverAllEfficiency"];

                        ws.Cells[row, 9].Value = rdr["Totaltime"];
                        ws.Cells[row, 10].Value = rdr["DownTime"];
                        ws.Cells[row, 11].Value = rdr["UtilisedTime"];
                        ws.Cells[row, 12].Value = rdr["MachineID1"];
                        ws.Cells[row, 12].Style.Font.Bold = true;
                        ws.Cells[row, 13].Value = rdr["Components1"];
                        ws.Cells[row, 14].Value = rdr["RejCount1"];
                        ws.Cells[row, 15].Value = rdr["PE"];
                        ws.Cells[row, 16].Value = rdr["AE"];
                        ws.Cells[row, 17].Value = rdr["QE"];
                        ws.Cells[row, 18].Value = rdr["OE"];
                        ws.Cells[row, 19].Value = rdr["AvgOE"];
                        ws.Cells[row, 20].Value = rdr["TotalTime1"];
                        ws.Cells[row, 21].Value = rdr["downtime1"];
                        ws.Cells[row, 22].Value = rdr["UtilisedTime1"];
                        row = row + 1;
                    }
                    ws.Cells[2, 1].Value = sMonth + " Month";

                    ws.Cells[2, 1, 2, 7].Style.Font.Bold = true;
                    ws.Cells[2, 1, 2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[2, 1, 2, 7].Style.Font.Size = 11;

                    ws.Cells[5, 8, row - 1, 8].Merge = true;
                    ws.Cells[5, 8, row - 1, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[5, 8, row - 1, 8].Style.Font.Size = 11;


                    ws.Cells[5, 19, row - 1, 19].Merge = true;
                    ws.Cells[5, 19, row - 1, 19].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[5, 19, row - 1, 19].Style.Font.Size = 11;


                    ws.Cells[2, 12].Value = PreviousMonth + " Month";
                    ws.Cells[2, 12, 2, 18].Style.Font.Bold = true;
                    ws.Cells[2, 12, 2, 18].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[2, 12, 2, 18].Style.Font.Size = 11;

                    ws.Cells[row, 1].Value = "Average OEE of " + sMonth + " Month";
                    ws.Cells[row, 1, row, 7].Merge = true;
                    ws.Cells[row, 1, row, 7].Style.Font.Bold = true;
                    ws.Cells[row, 1, row, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 1, row, 7].Style.Font.Size = 11;

                    ws.Cells[row, 8].Value = ws.Cells[row - 1, 8].Value;

                    ws.Cells[row, 12].Value = "Average OEE of " + PreviousMonth + " Month";
                    ws.Cells[row, 12, row, 18].Merge = true;
                    ws.Cells[row, 12, row, 18].Style.Font.Bold = true;
                    ws.Cells[row, 12, row, 18].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 12, row, 18].Style.Font.Size = 11;
                    ws.Cells[row, 19].Value = ws.Cells[row - 1, 19].Value;

                    CurrentMonthAvg = Convert.ToDouble(ws.Cells[row - 1, 8].Value);
                    PreMonthAvg = Convert.ToDouble(ws.Cells[row - 1, 19].Value);
                    ws.Cells[5, 24].Value = CurrentMonthAvg;
                    ws.Cells[6, 24].Value = PreMonthAvg;

                    using (ExcelRange range = ws.Cells[5, 1, row, 22])
                    {
                        range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    PlotGraphs(ws, 5, row - 1);
                    // ws.Protection.IsProtected = true;
                    //// ws.Protection.SetPassword("pctadmin$123");
                    // SetPrinterSettings(ws);
                    excelPackage.SaveAs(newFile);
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);

                }
                rdr.Close();

            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        public static void PlotGraphs(ExcelWorksheet wsDt, int startPos, int series)
        {
            var chart = (ExcelBarChart)wsDt.Drawings.AddChart("Monthwise OEE Report Report", eChartType.ColumnClustered);

            chart.SetPosition((series * 28), 20);

            chart.Legend.Remove();
            chart.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;

            chart.SetSize(710, 330);
            chart.Title.Text = "Monthwise OEE Report";

            var serie1 = chart.Series.Add(ExcelRange.GetAddress(5, 24, 6, 24), ExcelRange.GetAddress(5, 23, 6, 23));
        }

        //Vasavi Added
        private static void FillExcelSheet(string strtTime, string endTimez, string dst, string src, string machineId, string plantId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();

            try
            {
                FileInfo newFile = new FileInfo(dst);
                FileInfo tempFile = new FileInfo(src);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                string strPlantName = string.Empty;
                int row = 5, col = 1, start = row;
                string startTime = strtTime, endTime = endTimez;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetCockpitData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", startTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", "");
                cmd.Parameters.AddWithValue("@PlantID", "");
                strPlantName = string.Empty;

                List<decimal> PeAvgList = new List<decimal>();
                List<decimal> AeAvgList = new List<decimal>();
                List<decimal> QeAvgList = new List<decimal>();
                List<decimal> OeeAvgList = new List<decimal>();

                SqlDataReader rdr = cmd.ExecuteReader();


                if (rdr.HasRows)
                {
                    ws.Cells[2, 2].Value = startTime;
                    ws.Cells[2, 5].Value = endTime;
                    while (rdr.Read())
                    {

                        if (strPlantName != rdr["Remarks2"].ToString() && strPlantName != string.Empty)
                        {
                            col = 1;
                            ws.Cells[start, col].Value = "Average";
                            //Todo Vasavi
                            var cellPE = ws.Cells[start, col + 2];
                            var cellAE = ws.Cells[start, col + 3];
                            var cellQE = ws.Cells[start, col + 4];
                            var cellOEE = ws.Cells[start, col + 5];


                            cellPE.Formula = "AVERAGE(" + ws.Cells[row, 3].Address + ":" + ws.Cells[start - 1, 3].Address + ")";
                            cellAE.Formula = "AVERAGE(" + ws.Cells[row, 4].Address + ":" + ws.Cells[start - 1, 4].Address + ")";
                            cellQE.Formula = "AVERAGE(" + ws.Cells[row, 5].Address + ":" + ws.Cells[start - 1, 5].Address + ")";
                            cellOEE.Formula = "AVERAGE(" + ws.Cells[row, 6].Address + ":" + ws.Cells[start - 1, 6].Address + ")";
                            ws.Calculate();
                            if (!ExcelErrorValue.Values.IsErrorValue(cellPE.Value) && cellPE.Value.ToString() != "0") PeAvgList.Add(decimal.Parse(cellPE.Value.ToString()));
                            if (!ExcelErrorValue.Values.IsErrorValue(cellAE.Value) && cellAE.Value.ToString() != "0") AeAvgList.Add(decimal.Parse(cellAE.Value.ToString()));
                            if (!ExcelErrorValue.Values.IsErrorValue(cellQE.Value) && cellQE.Value.ToString() != "0") QeAvgList.Add(decimal.Parse(cellQE.Value.ToString()));
                            if (!ExcelErrorValue.Values.IsErrorValue(cellOEE.Value) && cellOEE.Value.ToString() != "0") OeeAvgList.Add(decimal.Parse(cellOEE.Value.ToString()));
                            // +		cellPE.Value	{#VALUE!}	object {OfficeOpenXml.ExcelErrorValue}


                            using (var range = ws.Cells[start, 1, start, 12])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                                range.Style.ShrinkToFit = false;
                            }

                            row = start + 1;
                            col = col + 1;
                            start = start + 1;
                        }

                        ws.Cells[start, 1].Value = rdr["Remarks2"].ToString();
                        ws.Cells[start, 2].Value = rdr["machineid"].ToString();

                        ws.Cells[start, 3].Value = decimal.Round(Convert.ToDecimal(rdr["ProductionEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 4].Value = decimal.Round(Convert.ToDecimal(rdr["AvailabilityEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 5].Value = decimal.Round(Convert.ToDecimal(rdr["QualityEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 6].Value = decimal.Round(Convert.ToDecimal(rdr["OverAllEfficiency"]), 2, MidpointRounding.AwayFromZero);

                        ws.Cells[start, 3].Value = ws.Cells[start, 3].Value.ToString() == "0" ? string.Empty : ws.Cells[start, 3].Value;
                        ws.Cells[start, 4].Value = ws.Cells[start, 4].Value.ToString() == "0" ? string.Empty : ws.Cells[start, 4].Value;
                        ws.Cells[start, 5].Value = ws.Cells[start, 5].Value.ToString() == "0" ? string.Empty : ws.Cells[start, 5].Value;
                        ws.Cells[start, 6].Value = ws.Cells[start, 6].Value.ToString() == "0" ? string.Empty : ws.Cells[start, 6].Value;



                        ws.Cells[start, 7].Value = Convert.ToInt32(rdr["Components"]);
                        ws.Cells[start, 8].Value = Convert.ToInt32(rdr["RejCount"]);
                        ws.Cells[start, 9].Value = decimal.Round(Convert.ToDecimal(rdr["ReturnPerHour"]), 2);
                        ws.Cells[start, 10].Value = (rdr["StrUtilisedTime"]);
                        ws.Cells[start, 11].Value = rdr["DownTime"].ToString();
                        ws.Cells[start, 12].Value = rdr["ManagementLoss"].ToString();
                        start = start + 1;
                        strPlantName = rdr["Remarks2"].ToString();
                    }
                }
                rdr.Close();
                col = 1;
                ws.Cells[start, col].Value = "Average";
                var cellP = ws.Cells[start, col + 2];
                var cellA = ws.Cells[start, col + 3];
                var cellQ = ws.Cells[start, col + 4];
                var cellOE = ws.Cells[start, col + 5];

                cellP.Formula = "AVERAGE(" + ws.Cells[row, 3].Address + ":" + ws.Cells[start - 1, 3].Address + ")";
                cellA.Formula = "AVERAGE(" + ws.Cells[row, 4].Address + ":" + ws.Cells[start - 1, 4].Address + ")";
                cellQ.Formula = "AVERAGE(" + ws.Cells[row, 5].Address + ":" + ws.Cells[start - 1, 5].Address + ")";
                cellOE.Formula = "AVERAGE(" + ws.Cells[row, 6].Address + ":" + ws.Cells[start - 1, 6].Address + ")";
                ws.Calculate();
                if (!ExcelErrorValue.Values.IsErrorValue(cellP.Value) && cellP.Value.ToString() != "0") PeAvgList.Add(decimal.Parse(cellP.Value.ToString()));
                if (!ExcelErrorValue.Values.IsErrorValue(cellA.Value) && cellA.Value.ToString() != "0") AeAvgList.Add(decimal.Parse(cellA.Value.ToString()));
                if (!ExcelErrorValue.Values.IsErrorValue(cellQ.Value) && cellQ.Value.ToString() != "0") QeAvgList.Add(decimal.Parse(cellQ.Value.ToString()));
                if (!ExcelErrorValue.Values.IsErrorValue(cellOE.Value) && cellOE.Value.ToString() != "0") OeeAvgList.Add(decimal.Parse(cellOE.Value.ToString()));

                start = start + 1;
                ws.Cells[start, col].Value = "Total Average";

                ws.Cells[start, col + 2].Value = AvgAEOrPEOrQEOrOEEE(PeAvgList);
                ws.Cells[start, col + 3].Value = AvgAEOrPEOrQEOrOEEE(AeAvgList);
                ws.Cells[start, col + 4].Value = AvgAEOrPEOrQEOrOEEE(QeAvgList);
                ws.Cells[start, col + 5].Value = AvgAEOrPEOrQEOrOEEE(OeeAvgList);

                using (var range = ws.Cells[start - 1, 1, start, 12])
                {
                    range.Style.Font.Bold = true;
                    //range.Style.Fill.PatternType = ExcelFillStyle.LightGray;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                    range.Style.ShrinkToFit = false;
                }



                using (ExcelRange range = ws.Cells[5, 1, start, 12])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                }

                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword("pctadmin$123");
                SetPrinterSettings(ws);
                excelPackage.SaveAs(newFile);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);


            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }

        }

        //Vasavi Added
        private static void FillExcelSheetForProductionCountBySlno(string strtTime, string endTimez, string dst, string src, string machineId, string plantId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                FileInfo newFile = new FileInfo(dst);
                FileInfo tempFile = new FileInfo(src);
                if (File.Exists(src))
                {
                    var di = new DirectoryInfo(src);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                }
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                string strPlantName = string.Empty;
                int row = 9, col = 1, start = row;
                string startTime = strtTime, endTime = endTimez;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cd = new SqlCommand(@"[s_GetShiftwiseRejectionReworkDetails]", sqlConn);
                cd.CommandType = System.Data.CommandType.StoredProcedure;
                cd.CommandTimeout = 600;
                cd.Parameters.AddWithValue("@StartDate", startTime);
                cd.Parameters.AddWithValue("@EndDate", startTime);
                cd.Parameters.AddWithValue("@ShiftIn", "");
                cd.Parameters.AddWithValue("@PlantID", plantId);
                cd.Parameters.AddWithValue("@MachineID", machineId);
                cd.Parameters.AddWithValue("@Param", "");
                string status = string.Empty;
                string strmachineid = string.Empty;
                string strDate = string.Empty;
                string strshift = string.Empty;
                string strTotalQuantity = string.Empty;
                string strTotalAccept = string.Empty;
                string strTotalReject = string.Empty;
                string strTotalRework = string.Empty;
                string strShiftWiseQuantity = string.Empty;
                string strShiftWiseAccept = string.Empty;
                string strShiftWiseReject = string.Empty;
                string strShiftWiseRework = string.Empty;

                SqlDataReader rs = cd.ExecuteReader();
                int count = 0;
                if (rs.HasRows)
                {
                    ws.Cells[4, 2].Value = startTime;
                    ws.Cells[4, 5].Value = endTime;

                    ws.Cells[4, 7].Value = plantId == string.Empty ? "ALL" : plantId;
                    ws.Cells[4, 9].Value = machineId == string.Empty ? "ALL" : machineId;

                    while (rs.Read())
                    {
                        if (count == 0)
                        {
                            strmachineid = rs["Machine"].ToString();
                            strDate = rs["ShDate"].ToString();
                            strshift = rs["ShftName"].ToString();
                            strTotalQuantity = rs["TotalQuantity"].ToString();
                            strTotalAccept = rs["TotalAccept"].ToString();
                            strTotalReject = rs["TotalRejection"].ToString();
                            strTotalRework = rs["TotalRework"].ToString();
                            strShiftWiseQuantity = rs["ShiftQuantity"].ToString();
                            strShiftWiseAccept = rs["ShiftAccept"].ToString();
                            strShiftWiseReject = rs["ShiftRejection"].ToString();
                            strShiftWiseRework = rs["ShiftRework"].ToString();
                        }
                        count++;

                        if (strmachineid != rs["Machine"].ToString() || strshift != rs["ShftName"].ToString() || strDate != rs["ShDate"].ToString())
                        {
                            ws.Cells[row, 1].Value = "Shiftwise Summary";
                            ws.Cells[row, 2].Value = "Shift" + "-" + strshift + " Total Quantity";
                            ws.Cells[row, 3].Value = strShiftWiseQuantity;
                            ws.Cells[row, 4].Value = "ShiftWise Accept";
                            ws.Cells[row, 5].Value = strShiftWiseAccept;
                            ws.Cells[row, 6].Value = "ShiftWise Rejection";
                            ws.Cells[row, 7].Value = strShiftWiseReject;
                            ws.Cells[row, 8].Value = "ShiftWise Rework";
                            ws.Cells[row, 9].Value = strShiftWiseRework;



                            using (var range = ws.Cells[row, 1, row, 9])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                                range.Style.ShrinkToFit = false;
                            }
                            row = row + 1;

                            if (strmachineid != rs["Machine"].ToString() || strDate != rs["ShDate"].ToString())
                            {
                                ws.Cells[row, 1].Value = "Daywise Summary";
                                ws.Cells[row, 2].Value = "Total Quantity";
                                ws.Cells[row, 3].Value = strTotalQuantity;
                                ws.Cells[row, 4].Value = "Total Accept";
                                ws.Cells[row, 5].Value = strTotalAccept;
                                ws.Cells[row, 6].Value = "Total Rejection";
                                ws.Cells[row, 7].Value = strTotalReject;
                                ws.Cells[row, 8].Value = "Total Rework";
                                ws.Cells[row, 9].Value = strTotalRework;

                                //  wrksht.Range(wrksht.Cells(row, 1), wrksht.Cells(row, 9)).interior.Color = RGB(250, 192, 144)
                                using (var range = ws.Cells[row, 1, row, 9])
                                {
                                    range.Style.Font.Bold = true;
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                                    range.Style.ShrinkToFit = false;
                                }
                                row = row + 1;

                                strmachineid = rs["Machine"].ToString();
                                strDate = rs["ShDate"].ToString();
                                strshift = rs["ShftName"].ToString();

                                strTotalQuantity = rs["TotalQuantity"].ToString();
                                strTotalAccept = rs["TotalAccept"].ToString();
                                strTotalReject = rs["TotalRejection"].ToString();
                                strTotalRework = rs["TotalRework"].ToString();
                            }

                            strshift = rs["ShftName"].ToString();
                            strmachineid = rs["Machine"].ToString();
                            strDate = rs["ShDate"].ToString();


                            strShiftWiseQuantity = rs["ShiftQuantity"].ToString();
                            strShiftWiseAccept = rs["ShiftAccept"].ToString();
                            strShiftWiseReject = rs["ShiftRejection"].ToString();
                            strShiftWiseRework = rs["ShiftRework"].ToString();
                        }


                        col = 1;
                        if (rs["Status"].ToString() == "Rejected")
                        {
                            using (var range = ws.Cells[row, 1, row, 9])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Red);// PatternType = ExcelFillStyle.DarkGrid; 
                                range.Style.Font.Color.SetColor(Color.White);
                                range.Style.ShrinkToFit = false;
                            }
                        }


                        ws.Cells[row, 1].Value = rs["Machine"].ToString();
                        ws.Cells[row, 2].Value = rs["Sttime"].ToString();
                        ws.Cells[row, 3].Value = rs["ShftName"].ToString();
                        ws.Cells[row, 4].Value = rs["OperatorID"].ToString();
                        ws.Cells[row, 5].Value = rs["Operator"].ToString();
                        ws.Cells[row, 6].Value = rs["Component"].ToString();
                        ws.Cells[row, 7].Value = rs["Operation"].ToString();
                        ws.Cells[row, 8].Value = rs["WorkOrderNumber"].ToString();
                        ws.Cells[row, 9].Value = rs["Status"].ToString();

                        row = row + 1;

                    }
                }
                rs.Close();

                int r = row;
                ws.Cells[row, 1].Value = "Shiftwise Summary";
                ws.Cells[row, 2].Value = "Shift - " + strshift + " Total Quantity";
                ws.Cells[row, 3].Value = strShiftWiseQuantity;
                ws.Cells[row, 4].Value = "ShiftWise Accept";
                ws.Cells[row, 5].Value = strShiftWiseAccept;
                ws.Cells[row, 6].Value = "ShiftWise Rejection";
                ws.Cells[row, 7].Value = strShiftWiseReject;
                ws.Cells[row, 8].Value = "ShiftWise Rework";
                ws.Cells[row, 9].Value = strShiftWiseRework;

                using (var range = ws.Cells[r, 1, r, 9])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                    range.Style.ShrinkToFit = false;
                }
                row = row + 1;

                ws.Cells[row, 1].Value = "Daywise Summary";
                ws.Cells[row, 2].Value = "Total Quantity";
                ws.Cells[row, 3].Value = strTotalQuantity;
                ws.Cells[row, 4].Value = "Total Accept";
                ws.Cells[row, 5].Value = strTotalAccept;
                ws.Cells[row, 6].Value = "Total Rejection";
                ws.Cells[row, 7].Value = strTotalReject;
                ws.Cells[row, 8].Value = "Total Rework";
                ws.Cells[row, 9].Value = strTotalRework;


                using (var range = ws.Cells[r + 1, 1, r + 1, 9])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);// Pat
                    range.Style.ShrinkToFit = false;
                }
                row = row + 1;

                using (ExcelRange range = ws.Cells[9, 1, row - 1, 9])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                }
                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword("pctadmin$123");
                SetPrinterSettings(ws);
                excelPackage.SaveAs(newFile);
                if (count > 0)
                {
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);
                }

            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }

        }

        private static void FillExportMachinewiseAlarmReport(string strtTime, string endTimez, string dst, string src, string machineId, string plantId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                string strPlantName = string.Empty;
                int row = 6, col = 1;
                string startTime = strtTime, endTime = endTimez;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetMachineAlarm]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", startTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", machineId);
                cmd.Parameters.AddWithValue("@AlarmCategory", "");
                cmd.Parameters.AddWithValue("@AlarmNumber", "");
                cmd.Parameters.AddWithValue("@PlantID", plantId);
                cmd.Parameters.AddWithValue("@param", "");


                strPlantName = string.Empty;
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    FileInfo newFile = new FileInfo(dst);
                    ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                    ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                    ws.Cells[4, 3].Value = startTime;
                    ws.Cells[4, 5].Value = endTime;
                    ws.Cells[3, 5].Value = plantId;
                    while (rdr.Read())
                    {
                        ws.Cells[row, 1].Value = rdr["SerialNo"];
                        ws.Cells[row, 2].Value = rdr["MachineID"];
                        ws.Cells[row, 3].Value = rdr["AlarmCategory"];
                        ws.Cells[row, 4].Value = rdr["AlarmNumber"];
                        ws.Cells[row, 5].Value = rdr["AlarmTime"];
                        ws.Cells[row, 6].Value = rdr["AlarmDescription"];
                        row = row + 1;

                    }
                    using (ExcelRange range = ws.Cells[6, 1, row - 1, 6])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    ws.Protection.IsProtected = true;
                    ws.Protection.SetPassword("pctadmin$123");
                    SetPrinterSettings(ws);
                    excelPackage.SaveAs(newFile);
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);

                }
                rdr.Close();

            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        private static void FillExportProductionAndRejectionReport(string strtTime, string endTimez, string Shift, string dst, string src, string machineId, string plantId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName, int ShiftID)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                string shiftIds = string.Empty;
                string param = string.Empty;
                shiftIds = "'" + Shift + "'";
                if (ShiftID == 3)
                {
                    shiftIds = "";
                    param = "schservice";
                }
                else
                {
                    param = "";
                }

                string mcStrMac = string.Empty;
                string ComponentName = string.Empty;
                string Shiftname = string.Empty;
                string OperatorName = string.Empty;
                string mcStr = string.Empty;
                string daywiseAE = string.Empty, daywisePE = string.Empty, daywiseQE = string.Empty, daywiseOEE = string.Empty, prevday = string.Empty,
                    daywiseRejection = string.Empty, daywiseDowntime = string.Empty;

                int row = 2, rowstart = 2, CountStartRow, column = 2, strcolumn = 2, count = 0, rownum = 3, rowmerge = 3;

                string startTime = strtTime, endTime = endTimez;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetWiproSantoor_ProdDownReport]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                int i = 0;
                cmd.Parameters.AddWithValue("@StartDate", strtTime);
                cmd.Parameters.AddWithValue("@Enddate", endTimez);
                cmd.Parameters.AddWithValue("@ShiftIn", shiftIds);
                string[] values = Shift.Split(',');
                cmd.Parameters.AddWithValue("@PlantID", plantId);
                cmd.Parameters.AddWithValue("@MachineID", machineId);
                cmd.Parameters.AddWithValue("@RptProd_down", "ProductionRejection");
                cmd.Parameters.AddWithValue("@Param", param);
                SqlDataReader RecordSet = cmd.ExecuteReader();

                #region MyRegion
                if (RecordSet.HasRows)
                {
                    FileInfo newFile = new FileInfo(dst);
                    ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                    ExcelWorksheet wrksht = excelPackage.Workbook.Worksheets[1];
                    // Logger.WriteDebugLog("created excel sheet");
                    int rowcount = GetLastUsedRow(wrksht);
                    rowcount = rowcount + 1;
                    string strMachine = String.Empty; string StrDate = String.Empty; string StrShift = String.Empty;
                    string StrComponent = string.Empty; string StrOperator = string.Empty; string DaywiseRejQty = string.Empty;
                    string DaywiseDowntime = string.Empty; string DaywiseAE = string.Empty;
                    string DaywisePE = string.Empty; string DaywiseOEE = string.Empty;
                    string DaywiseQE = string.Empty;


                    string h1Actual = string.Empty; string h1Down = string.Empty;
                    string h2Actual = string.Empty; string h2Down = string.Empty;
                    string h3Actual = string.Empty; string h3Down = string.Empty;
                    string h4Actual = string.Empty; string h4Down = string.Empty;
                    string h5Actual = string.Empty; string h5Down = string.Empty;
                    string h6Actual = string.Empty; string h6Down = string.Empty;
                    string h7Actual = string.Empty; string h7Down = string.Empty;
                    string h8Actual = string.Empty; string h8Down = string.Empty;


                    row = rowcount;
                    column = 1;


                    rowstart = row;
                    CountStartRow = row;
                    while (RecordSet.Read())
                    {
                        // Logger.WriteDebugLog("inside while ");
                        if (count == 0)
                        {
                            strMachine = RecordSet["MachineID"].ToString();
                            StrDate = RecordSet["Date"].ToString();
                            StrShift = RecordSet["ShiftName"].ToString();
                            StrComponent = RecordSet["ComponentID"].ToString();
                            StrOperator = RecordSet["OperatorID"].ToString();
                            DaywiseRejQty = RecordSet["DaywiseRejQty"].ToString();
                            DaywiseDowntime = RecordSet["DaywiseDowntime"].ToString();
                            DaywiseAE = RecordSet["DaywiseAE"].ToString();
                            DaywisePE = RecordSet["DaywisePE"].ToString();
                            DaywiseQE = RecordSet["DaywiseQE"].ToString();
                            DaywiseOEE = RecordSet["DaywiseOEE"].ToString();
                            h1Actual = RecordSet["Hour1Actual"].ToString();
                            h2Actual = RecordSet["Hour2Actual"].ToString();
                            h3Actual = RecordSet["Hour3Actual"].ToString();
                            h4Actual = RecordSet["Hour4Actual"].ToString();
                            h5Actual = RecordSet["Hour5Actual"].ToString();
                            h6Actual = RecordSet["Hour6Actual"].ToString();
                            h7Actual = RecordSet["Hour7Actual"].ToString();
                            h8Actual = RecordSet["Hour8Actual"].ToString();
                            h1Down = RecordSet["Hour1DT"].ToString();
                            h2Down = RecordSet["Hour2DT"].ToString();
                            h3Down = RecordSet["Hour3DT"].ToString();
                            h4Down = RecordSet["Hour4DT"].ToString();
                            h5Down = RecordSet["Hour5DT"].ToString();
                            h6Down = RecordSet["Hour6DT"].ToString();
                            h7Down = RecordSet["Hour7DT"].ToString();
                            h8Down = RecordSet["Hour8DT"].ToString();



                            wrksht.Cells[row, column].Value = RecordSet["Date"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["ShiftName"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["MachineID"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["ComponentID"].ToString();
                            column = column + 2;
                            wrksht.Cells[row, column].Value = RecordSet["OperatorID"].ToString();
                            column = column + 1;
                            strcolumn = column;


                            wrksht.Cells[row, 31].Value = Convert.ToDecimal(RecordSet["HourlyTarget"]);
                            wrksht.Cells[row, 32].Value = Convert.ToDecimal(RecordSet["ShiftTarget"]);
                            wrksht.Cells[row, 33].Value = RecordSet["ShftActualCount"];

                            wrksht.Cells[row, 34].Value = RecordSet["TotalRejection"];
                            wrksht.Cells[row, 35].Value = RecordSet["MachinewiseDowntime"];
                            wrksht.Cells[row, 36].Value = RecordSet["AE"];
                            wrksht.Cells[row, 37].Value = RecordSet["PE"];
                            wrksht.Cells[row, 38].Value = RecordSet["QE"];
                            wrksht.Cells[row, 39].Value = RecordSet["OEE"];
                            //Logger.WriteDebugLog("inside while  count = 0");
                        }
                        count++;
                        if (count == 1)
                        {
                            row = rowcount;
                            column = 1;
                            rowstart = row;
                            CountStartRow = row;
                        }
                        if (StrDate.ToString() != RecordSet["Date"].ToString() || StrShift != RecordSet["ShiftName"].ToString() || strMachine != RecordSet["MachineID"].ToString() || StrComponent != RecordSet["ComponentID"].ToString() || StrOperator != RecordSet["OperatorID"].ToString())
                        {
                            //Logger.WriteDebugLog("inside while  in if");
                            if ((StrDate.ToString() != RecordSet["Date"].ToString() || strMachine != RecordSet["MachineID"].ToString()) && ShiftID == 3)
                            {
                                row = row + 1;

                                wrksht.Cells[row, 2].Value = "Day";
                                wrksht.Cells[row, 1].Value = StrDate;


                                wrksht.Cells[row, 1, row, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                wrksht.Cells[row, 1, row, 40].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 224, 225));
                                wrksht.Cells[row, 7].Value = h1Actual;
                                wrksht.Cells[row, 8].Value = h1Down;
                                wrksht.Cells[row, 9].Value = h2Actual;
                                wrksht.Cells[row, 10].Value = h2Down;
                                wrksht.Cells[row, 11].Value = h3Actual;
                                wrksht.Cells[row, 12].Value = h3Down;
                                wrksht.Cells[row, 13].Value = h4Actual;
                                wrksht.Cells[row, 14].Value = h4Down;
                                wrksht.Cells[row, 15].Value = h5Actual;
                                wrksht.Cells[row, 16].Value = h5Down;
                                wrksht.Cells[row, 17].Value = h6Actual;
                                wrksht.Cells[row, 18].Value = h6Down;
                                wrksht.Cells[row, 19].Value = h7Actual;
                                wrksht.Cells[row, 20].Value = h7Down;
                                wrksht.Cells[row, 21].Value = h8Actual;
                                wrksht.Cells[row, 22].Value = h8Down;


                                wrksht.Cells[row, 34].Value = DaywiseRejQty;
                                wrksht.Cells[row, 35].Value = DaywiseDowntime;
                                wrksht.Cells[row, 36].Value = DaywiseAE;
                                wrksht.Cells[row, 37].Value = DaywisePE;
                                wrksht.Cells[row, 38].Value = DaywiseQE;
                                wrksht.Cells[row, 39].Value = DaywiseOEE;

                                DaywiseRejQty = RecordSet["DaywiseRejQty"].ToString();
                                DaywiseDowntime = RecordSet["DaywiseDowntime"].ToString();
                                DaywiseAE = RecordSet["DaywiseAE"].ToString();
                                DaywisePE = RecordSet["DaywisePE"].ToString();
                                DaywiseQE = RecordSet["DaywiseQE"].ToString();
                                DaywiseOEE = RecordSet["DaywiseOEE"].ToString();

                                h1Actual = RecordSet["Hour1Actual"].ToString();
                                h2Actual = RecordSet["Hour2Actual"].ToString();
                                h3Actual = RecordSet["Hour3Actual"].ToString();
                                h4Actual = RecordSet["Hour4Actual"].ToString();
                                h5Actual = RecordSet["Hour5Actual"].ToString();
                                h6Actual = RecordSet["Hour6Actual"].ToString();
                                h7Actual = RecordSet["Hour7Actual"].ToString();
                                h8Actual = RecordSet["Hour8Actual"].ToString();
                                h1Down = RecordSet["Hour1DT"].ToString();
                                h2Down = RecordSet["Hour2DT"].ToString();
                                h3Down = RecordSet["Hour3DT"].ToString();
                                h4Down = RecordSet["Hour4DT"].ToString();
                                h5Down = RecordSet["Hour5DT"].ToString();
                                h6Down = RecordSet["Hour6DT"].ToString();
                                h7Down = RecordSet["Hour7DT"].ToString();
                                h8Down = RecordSet["Hour8DT"].ToString();

                                wrksht.Cells[rowstart, 34, row - 1, 34].Merge = true;
                                wrksht.Cells[rowstart, 35, row - 1, 35].Merge = true;
                                wrksht.Cells[rowstart, 36, row - 1, 36].Merge = true;
                                wrksht.Cells[rowstart, 37, row - 1, 37].Merge = true;
                                wrksht.Cells[rowstart, 38, row - 1, 38].Merge = true;
                                rowstart = row + 1;
                                CountStartRow = row + 1;
                            }
                            row = row + 1;
                            column = 1;
                            wrksht.Cells[row, column].Value = RecordSet["Date"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["ShiftName"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["MachineID"].ToString();
                            column = column + 1;
                            wrksht.Cells[row, column].Value = RecordSet["ComponentID"].ToString();
                            column = column + 2;
                            wrksht.Cells[row, column].Value = RecordSet["OperatorID"].ToString();
                            column = column + 1;
                            strcolumn = column;

                            wrksht.Cells[row, strcolumn].Value = Convert.ToDecimal(RecordSet["Actual"].ToString());
                            strcolumn = strcolumn + 1;

                            wrksht.Cells[row, strcolumn].Value = Convert.ToDecimal(RecordSet["HourlyDowntime"].ToString());
                            strcolumn = strcolumn + 1;

                            wrksht.Cells[row, 31].Value = Convert.ToDecimal(RecordSet["HourlyTarget"]);
                            wrksht.Cells[row, 32].Value = Convert.ToDecimal(RecordSet["ShiftTarget"]);
                            wrksht.Cells[row, 33].Value = RecordSet["ShftActualCount"].ToString();



                            if (strMachine != RecordSet["MachineID"].ToString() || StrDate.ToString() != RecordSet["Date"].ToString() || StrShift != RecordSet["ShiftName"].ToString())
                            {
                                wrksht.Cells[row, 34].Value = RecordSet["TotalRejection"].ToString();
                                column = column + 1;
                                wrksht.Cells[row, 35].Value = RecordSet["MachinewiseDowntime"].ToString();
                                column = column + 1;
                                wrksht.Cells[row, 36].Value = RecordSet["AE"].ToString();
                                column = column + 1;
                                wrksht.Cells[row, 37].Value = RecordSet["PE"].ToString();
                                column = column + 1;
                                wrksht.Cells[row, 38].Value = RecordSet["QE"].ToString();
                                column = column + 1;
                                wrksht.Cells[row, 39].Value = RecordSet["OEE"].ToString();
                                column = column + 1;

                            }
                            if (strMachine == RecordSet["MachineID"].ToString() && StrDate.ToString() == RecordSet["Date"].ToString() && StrShift != RecordSet["ShiftName"].ToString())
                            {
                                wrksht.Cells[rowstart, 34, row - 1, 34].Merge = true;
                                wrksht.Cells[rowstart, 35, row - 1, 35].Merge = true;
                                wrksht.Cells[rowstart, 36, row - 1, 36].Merge = true;
                                wrksht.Cells[rowstart, 37, row - 1, 37].Merge = true;
                                wrksht.Cells[rowstart, 38, row - 1, 38].Merge = true;
                                wrksht.Cells[rowstart, 39, row - 1, 39].Merge = true;
                                rowstart = row;
                            }

                            strMachine = RecordSet["MachineID"].ToString();
                            StrDate = RecordSet["Date"].ToString();
                            StrShift = RecordSet["ShiftName"].ToString();
                            StrComponent = RecordSet["ComponentID"].ToString();
                            StrOperator = RecordSet["OperatorID"].ToString();
                        }
                        else
                        {

                            wrksht.Cells[row, strcolumn].Value = Convert.ToDecimal(RecordSet["Actual"].ToString());
                            strcolumn = strcolumn + 1;

                            wrksht.Cells[row, strcolumn].Value = Convert.ToDecimal(RecordSet["HourlyDowntime"].ToString());
                            strcolumn = strcolumn + 1;


                        }
                    }//wend
                    row = row + 1;
                    if (ShiftID == 3)
                    {
                        wrksht.Cells[row, 2].Value = "Day";
                        wrksht.Cells[row, 1].Value = StrDate;

                        wrksht.Cells[row, 1, row, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wrksht.Cells[row, 1, row, 40].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 224, 225));
                        wrksht.Cells[row, 7].Value = h1Actual;
                        wrksht.Cells[row, 8].Value = h1Down;
                        wrksht.Cells[row, 9].Value = h2Actual;
                        wrksht.Cells[row, 10].Value = h2Down;
                        wrksht.Cells[row, 11].Value = h3Actual;
                        wrksht.Cells[row, 12].Value = h3Down;
                        wrksht.Cells[row, 13].Value = h4Actual;
                        wrksht.Cells[row, 14].Value = h4Down;
                        wrksht.Cells[row, 15].Value = h5Actual;
                        wrksht.Cells[row, 16].Value = h5Down;
                        wrksht.Cells[row, 17].Value = h6Actual;
                        wrksht.Cells[row, 18].Value = h6Down;
                        wrksht.Cells[row, 19].Value = h7Actual;
                        wrksht.Cells[row, 20].Value = h7Down;
                        wrksht.Cells[row, 21].Value = h8Actual;
                        wrksht.Cells[row, 22].Value = h8Down;
                        wrksht.Cells[row, 34].Value = DaywiseRejQty;
                        wrksht.Cells[row, 35].Value = DaywiseDowntime;
                        wrksht.Cells[row, 36].Value = DaywiseAE;
                        wrksht.Cells[row, 37].Value = DaywisePE;
                        wrksht.Cells[row, 38].Value = DaywiseQE;
                        wrksht.Cells[row, 39].Value = DaywiseOEE;
                    }
                    wrksht.Cells[rowstart, 34, row - 1, 34].Merge = true;
                    wrksht.Cells[rowstart, 35, row - 1, 35].Merge = true;
                    wrksht.Cells[rowstart, 36, row - 1, 36].Merge = true;
                    wrksht.Cells[rowstart, 37, row - 1, 37].Merge = true;
                    wrksht.Cells[rowstart, 38, row - 1, 38].Merge = true;
                    wrksht.Cells[rowstart, 39, row - 1, 39].Merge = true;
                    wrksht.Cells.AutoFitColumns();
                    SetPrinterSettings(wrksht);

                    wrksht.Column(23).Hidden = true;
                    wrksht.Column(24).Hidden = true;
                    wrksht.Column(25).Hidden = true;
                    wrksht.Column(26).Hidden = true;
                    wrksht.Column(27).Hidden = true;
                    wrksht.Column(28).Hidden = true;
                    wrksht.Column(29).Hidden = true;
                    wrksht.Column(30).Hidden = true;
                    //Logger.WriteDebugLog("before save excel= 0");
                    excelPackage.SaveAs(newFile);
                    // Logger.WriteDebugLog("After saving  excel");

                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);
                    //  Logger.WriteDebugLog("After sending  mail");
                }
                else
                {
                    Logger.WriteDebugLog("No data found to export. email not required....");
                }
                RecordSet.Close();

                #endregion

            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Production and Rejection Report Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        private static void SetPrinterSettings(ExcelWorksheet wsDt)
        {
            wsDt.PrinterSettings.Orientation = eOrientation.Landscape;
            wsDt.PrinterSettings.PaperSize = ePaperSize.A4;
            wsDt.PrinterSettings.LeftMargin = new decimal(.25);
            wsDt.PrinterSettings.RightMargin = new decimal(.25);
            wsDt.PrinterSettings.TopMargin = new decimal(.25);
            wsDt.PrinterSettings.BottomMargin = new decimal(.25);
            wsDt.PrinterSettings.PrintArea = wsDt.Cells[wsDt.Dimension.Start.Address + ":" + wsDt.Dimension.End.Address];
        }

        public static decimal AvgAEOrPEOrQEOrOEEE(List<decimal> listings)
        {
            decimal total = 0;
            foreach (decimal p in listings)
            {
                total += p;
            }

            decimal avg = 0.00m;

            if (listings.Count != 0)
            {
                avg = total / listings.Count;
            }
            return avg;
        }

        private static void FillExcelSheetData(string strtTime, string endTimez, string dst, string src, string machineId, string plantId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();

            try
            {
                FileInfo newFile = new FileInfo(dst);
                FileInfo tempFile = new FileInfo(src);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                string strMacName = string.Empty;
                int row = 5, col = 0, start = row;
                string startTime = strtTime, endTime = endTimez;
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetEfficiency]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 1800;
                cmd.Parameters.AddWithValue("@Starttime", startTime);
                cmd.Parameters.AddWithValue("@TimeAxis", "");
                cmd.Parameters.AddWithValue("@Type", "Console");
                cmd.Parameters.AddWithValue("@ComparisonParam", "OEE");
                cmd.Parameters.AddWithValue("@MachineID", "");
                cmd.Parameters.AddWithValue("@PlantID", "");
                cmd.Parameters.AddWithValue("@ShiftName", "");

                SqlDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        if (strMacName != rdr["machineid"].ToString())
                        {
                            col = 1;
                            row = row + 1;
                            ws.Cells[row, col].Value = rdr["machineid"].ToString();
                            col = col + 1;
                            start = start + 1;
                        }

                        if (((rdr["shftnm"]).ToString()).Equals("A"))
                        {
                            ws.Cells[5, 2].Value = rdr["operator"].ToString();
                        }
                        else if (((rdr["shftnm"]).ToString()).Equals("B"))
                        {
                            ws.Cells[5, 3].Value = rdr["operator"].ToString();
                        }
                        else
                        {
                            ws.Cells[5, 4].Value = rdr["operator"].ToString();
                        }

                        ws.Cells[start, col].Value = Convert.ToInt32(rdr["Components"]);
                        ws.Cells[start, col + 6].Value = decimal.Round(Convert.ToDecimal(rdr["PE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, col + 11].Value = decimal.Round(Convert.ToDecimal(rdr["AE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, col + 16].Value = decimal.Round(Convert.ToDecimal(rdr["OE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, col + 21].Value = decimal.Round(Convert.ToDecimal(rdr["QualityEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, col + 26].Value = rdr["downtime"];
                        col = col + 1;
                        strMacName = rdr["machineid"].ToString();
                    }
                }

                rdr.NextResult();

                if (rdr.HasRows)
                {
                    start = 6;
                    while (rdr.Read())
                    {
                        ws.Cells[start, 5].Value = Convert.ToInt32(rdr["Components"]);
                        ws.Cells[start, 11].Value = decimal.Round(Convert.ToDecimal(rdr["PE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 16].Value = decimal.Round(Convert.ToDecimal(rdr["AE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 21].Value = decimal.Round(Convert.ToDecimal(rdr["OE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 26].Value = decimal.Round(Convert.ToDecimal(rdr["QualityEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 31].Value = rdr["downtime"];
                        ws.Cells[start, 33].Value = rdr["TurnOver"];
                        start = start + 1;
                    }
                }

                rdr.NextResult();

                if (rdr.HasRows)
                {
                    start = 6;
                    while (rdr.Read())
                    {
                        ws.Cells[start, 6].Value = Convert.ToInt32(rdr["Components"]);
                        ws.Cells[start, 7].Value = rdr["RejCount"];
                        ws.Cells[start, 12].Value = decimal.Round(Convert.ToDecimal(rdr["PE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 17].Value = decimal.Round(Convert.ToDecimal(rdr["AE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 22].Value = decimal.Round(Convert.ToDecimal(rdr["OE"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 27].Value = decimal.Round(Convert.ToDecimal(rdr["QualityEfficiency"]), 2, MidpointRounding.AwayFromZero);
                        ws.Cells[start, 32].Value = rdr["downtime"];
                        ws.Cells[start, 34].Value = rdr["TurnOver"];
                        ws.Cells[start, 35].Value = decimal.Round(Convert.ToDecimal(rdr["MonthRevenueLoss"]), 2, MidpointRounding.AwayFromZero);
                        start = start + 1;
                    }
                }

                rdr.Close();

                //GetWindowThreadProcessId(excelPackage.Hwnd, out pid);
                ws.Cells["C2"].Value = string.Format("{0:dd-MMM-yy}", DateTime.Parse(strtTime));
                using (ExcelRange range = ws.Cells[6, 1, start - 1, 35])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                }
                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword("pctadmin$123");
                SetPrinterSettings(ws);

                excelPackage.SaveAs(newFile);
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, exportFileName);

            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }


        public static int GetLastUsedRow(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                row--;
            }
            return row;
        }



        public static bool ExportProductionReportMachinewise(string dst, string ExportPath, string ExportedReportFile, int ExportType, int DayBefores, string Shift, string MachineId, string operators, string sttime, string ndtime, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string CompanyName, bool MachineAE)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlDataReader sdr = null;
            string destPath = string.Empty;
            if (!File.Exists(dst))
            {
                Logger.WriteDebugLog("Template is not found on " + dst);
                return false;
            }
            if (!Directory.Exists(ExportPath))
            {
                Directory.CreateDirectory(ExportPath);
            }
            destPath = Path.Combine(ExportPath, string.Format("SM_MachineWiseProductionDetails_{0}.xlsx", DateTime.Now.ToString("dd_MMM_yyyy_HH_mm")));
            if (File.Exists(destPath))
            {
               var dir = new DirectoryInfo(destPath);
                dir.Attributes &= ~FileAttributes.ReadOnly;
                File.Delete(destPath);
            }
            File.Copy(dst, destPath, true);
            try
            {
                SqlCommand cmd = new SqlCommand(@"[dbo].[S_readMachinewiseProductiondetails]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@StartDate", sttime);
                cmd.Parameters.AddWithValue("@Enddate", ndtime);
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@ShiftIn", Shift);
                cmd.Parameters.AddWithValue("@Param", "Shiftwise");
                sdr = cmd.ExecuteReader();
                FileInfo newFile = new FileInfo(destPath);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                if (sdr.HasRows)
                {
                    ws.Cells[4, 2].Value = Convert.ToDateTime(sttime).ToString("dd-MMM-yyyy");
                    ws.Cells[4, 4].Value = Convert.ToDateTime(ndtime).ToString("dd-MMM-yyyy");
                    ws.Cells[4, 6].Value = plantid;
                    ws.Cells[4, 8].Value = MachineId;
                    int row = 7;
                    while (sdr.Read())
                    {

                        if (!string.IsNullOrEmpty(sdr["Udate"].ToString()) && !Convert.IsDBNull(sdr["Udate"]))
                        {
                            ws.Cells[row, 1].Value = Convert.ToDateTime(sdr["Udate"].ToString()).ToString("dd-MMM-yyyy");
                        }
                        ws.Cells[row, 2].Value = sdr["UShift"].ToString();
                        ws.Cells[row, 3].Value = sdr["Machineid"].ToString();
                        ws.Cells[row, 4].Value = sdr["UtilisedTime"].ToString();
                        ws.Cells[row, 5].Value = sdr["DownTime"].ToString();
                        ws.Cells[row, 6].Value = sdr["ManagementLoss"];
                        ws.Cells[row, 7].Value = sdr["PDT"];
                        ws.Cells[row, 8].Value = sdr["QTY"];
                        ws.Cells[row, 9].Value = sdr["PE"];
                        ws.Cells[row, 10].Value = sdr["AE"];
                        ws.Cells[row, 11].Value = sdr["OEE"];
                        row++;
                    }
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Shiftwise Production Report Machinewise generated successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, destPath, Path.GetFileName(destPath));
                }
                else
                {
                    Logger.WriteDebugLog("No Data found for Shiftwise Production Report(Time-Consolidated): sttime = " + sttime);
                }

            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Production Report Machinewise Excel File..!!\n " + ex.Message);
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
                if (sdr != null) sdr.Close();
            }
            return true;
        }

        public static void SendFileShareFiles(bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {
            var _directoryPath = ConfigurationManager.AppSettings["FileShareFolderPath"].ToString();
            var _password = ConfigurationManager.AppSettings["Password_Fileshare"].ToString();
            var _userName = ConfigurationManager.AppSettings["UserID_Fileshare"].ToString();

            var downloadFolder = Path.Combine(_appPath, "DownLoadedFile");
            if (!Directory.Exists(downloadFolder))
            {
                Directory.CreateDirectory(downloadFolder);
            }

            NetworkConnection nc = null;
            string File_NameFolder = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(_userName) && !string.IsNullOrEmpty(_password))
                {
                    try
                    {
                        nc = new NetworkConnection(_directoryPath, new NetworkCredential(_userName, _password));
                    }
                    catch (Exception exx)
                    {
                        Logger.WriteErrorLog(exx.ToString());
                    }
                }
                var AllFiles = Directory.GetFiles(_directoryPath);
                if (AllFiles.Length > 0)
                {
                    var fileInfo = AllFiles.Where(f => Path.GetExtension(f).ToLower() == ".xlsx" || 
                    Path.GetExtension(f).ToLower() == ".xls").Select(f => new FileInfo(f))
                                    .OrderByDescending(fi => fi.LastWriteTime)
                                    .FirstOrDefault();

                    if (fileInfo != null)
                    {
                        string file = fileInfo.FullName;
                        File_NameFolder = Path.Combine(downloadFolder, Path.GetFileName(file));

                        if (Path.GetExtension(file).Equals(".XLS", StringComparison.OrdinalIgnoreCase) ||
                        Path.GetExtension(file).Equals(".XLSX", StringComparison.OrdinalIgnoreCase))
                        {
                            if (File.Exists(File_NameFolder) == true)
                            {
                                File.Delete(File_NameFolder);
                                File.Copy(file, File_NameFolder);
                            }
                            else
                            {
                                File.Copy(file, File_NameFolder);
                            }
                        }
                    }
                }
                else
                {
                    Logger.WriteDebugLog("No Sand Report related file found in " + _directoryPath);
                }
                if (nc != null) nc.Dispose();
                if (!string.IsNullOrEmpty(File_NameFolder) && File.Exists(File_NameFolder))
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, File_NameFolder, Path.GetFileName(File_NameFolder));

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
        }

        public static bool FloatSheetGenerateReport(string strReportFile, string ExportPath, string ExportedReportFile,
    string sttime, bool Email_Flag, string Email_List_To, string Email_List_CC)
        {
            string dst = string.Empty, src = string.Empty;
            SqlConnection sqlConn = null;//new SqlConnection("Data Source=PCT-DEV8\\SQLEXP2012; Initial Catalog=SmartMachine;User ID=sa; password=pctadmin$123");
                                         // sqlConn = new SqlConnection(@"Data Source=PCT-DEV11\SQLEXP2012; Initial Catalog=Bosch_Nashik;User ID=sa; password=pctadmin$123");
            try
            {
                string SDate = string.Format("{0:yyyy-MMM-dd hh:mm:ss tt}", sttime);

                string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                src = APath + @"\Reports\Local MAE Float sheet 2017.xlsx";
                if (!File.Exists(src))
                {
                    Logger.WriteDebugLog("Template is not found on " + src);
                    return false;
                }
                dst = ExportPath + @"Local MAE Float sheet_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(SDate)) + ".xlsx";

                try
                {
                    File.Copy(src, dst, true);
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }

                if (!File.Exists(dst))
                {
                    return false;
                }

                //sqlConn.Open();
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[S_GetBoschNashik_FloatSheetData]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Date", string.Format("{0:yyyy-MM-dd hh:mm:ss}", sttime));//sttime.ToString("yyyy-MM-dd hh:mm:ss"))"2016-11-10";            
                cmd.Parameters.AddWithValue("@Machineid", "");
                cmd.Parameters.AddWithValue("@PlantID", "");
                cmd.Parameters.AddWithValue("@GroupID", "");//groupId);


                SqlDataReader rdr = cmd.ExecuteReader();
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage pck = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = pck.Workbook.Worksheets[2];
                ws.Cells["C2"].Value = sttime;
                var columns = new List<string>();
                int row = 5, col = 3;
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    columns.Add(rdr.GetName(i));
                }

                foreach (var item in columns)
                {
                    if (item != "Line")
                    {
                        ws.Cells[4, col].Value = item;
                        col++;
                    }
                }
                int lenth = columns.Count;
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        col = 2;
                        for (int i = 0; i < columns.Count; i++)
                        {
                            if (rdr["Line"].ToString() != "TOTAL")
                            {
                                ws.Cells[row, col].Value = rdr[i];
                                col = col + 1;
                            }
                        }
                        row = row + 1;
                    }
                }
                if (rdr != null)
                {
                    rdr.Close();
                }

                pck.SaveAs(newFile);

                Logger.WriteDebugLog("Report generated sucessfully.");
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, "", dst, ExportedReportFile);
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Excel File..!!\n " + ex.Message);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }

            return true;
        }

        internal static void ExportDailyRejectionReport(string strtTime, string endTime, string strReportFile, string ExportPath, string ExportedReportFile,
         string MachineId, string operators, string sttime,
         string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
         string Email_List_BCC)
        {
            //strtTime = "2017-Nov-12 06:00:00 AM"; //g: testdates
            //endTime = "2017-Nov-19 06:00:00 AM";
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = ExportPath + @"SM_DailyRejectionReport_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(strtTime)) + ".xlsx";
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FillDailyRejectionReportReport(strtTime, endTime, dst, strReportFile, MachineId, plantid, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, ExportedReportFile);
                Logger.WriteDebugLog("Data Exported successfully..!! \n View the Excel Sheet Data.");
            }

            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error..!! \n" + ex.Message);
            }
        }


        private static void FillDailyRejectionReportReport(string strtTime, string endTime, string dst, string strReportFile, string MachineId, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string ExportedReportFile)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }
                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }

                string strPlantName = string.Empty;
                string tempmachineid = string.Empty;

                List<string> lstMachineNames = new List<string>();
                List<string> lstMacTotal = new List<string>();
                List<string> lstMacFreqTotal = new List<string>();
                string mxkSectohhmmss = string.Empty;
                int r = 6;
                string startTime = strtTime;


                string DownID = string.Empty;
                string PrevDownFreq = string.Empty;
                string DownIDTotal = string.Empty;
                string downtime = string.Empty;

                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"[s_GetRejectioncodeDetails]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Starttime", strtTime);
                cmd.Parameters.AddWithValue("@EndTime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", MachineId);
                cmd.Parameters.AddWithValue("@PlantID", plantid);
                string PrevMachine = string.Empty;
                strPlantName = string.Empty;
                string Prevdown = string.Empty;
                SqlDataReader rdr = cmd.ExecuteReader();

                r = 6;

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Column(1).Style.Numberformat.Format = "yyyy-mm-dd";
                if (rdr.HasRows)
                {
                    ws.Cells[3, 2].Value = strtTime;
                    ws.Cells[3, 5].Value = endTime;
                    ws.Cells[3, 7].Value = plantid.Equals("") ? "All" : plantid;
                    ws.Cells[3, 9].Value = MachineId.Equals("") ? "All" : MachineId;
                    ws.Cells["A5"].LoadFromDataReader(rdr, true);

                    ws.Name = "DailyRejection";
                    excelPackage.SaveAs(newFile);
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile); //g: 
                }
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n While Creating Excel File..!!\n " + ex.Message + Environment.NewLine + ex.StackTrace);
                dst = string.Empty;
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        internal static void ExportHourlyMachinewiseProductionReport(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {
            //strtTime = "2018-01-12"; //g: testdates
            //endTime = "2018-01-12";
            string dst = string.Empty;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath, string.Format("SM_HourlyMachinewiseProductionReport_{0:ddMMMyyyy}.xlsx", DateTime.Parse(strtTime)));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);



                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }
                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }

                int r = 6, c = 1;
                bool dataAvailable = false;
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Name = "Production";

                ws.Column(1).Style.Numberformat.Format = "dd-mmm-yy";


                System.Data.DataTable dt = AccessReportData.GetShiftIDsandNames();


                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    ws.Cells[r + i, c].Value = dr["ShiftName"].ToString();
                }
                dt.Clear();

                string shift = "";
                string firstshift = "";
                SqlDataReader rdr = AccessReportData.GetHourlyMachinewiseProduction(strtTime, endTime, plantid, MachineId);
                dt.Load(rdr);
                rdr.Close();

                if (dt.Rows.Count > 0)
                {
                    //ws.Cells[17, 7, dt.Rows.Count + 17, 14].Style.Numberformat.Format = "0";
                    r = 6;
                    c = 2;
                    dataAvailable = true;
                    ws.Cells[3, 1].Value = strtTime;
                    ws.Cells[3, 2].Value = endTime;
                    ws.Cells[3, 3].Value = plantid.Equals("") ? "All" : plantid;
                    ws.Cells[3, 4].Value = MachineId.Equals("") ? "All" : MachineId;
                    ws.Cells[3, 5].Value = "All";

                    int i = 0;
                    int n = dt.Rows.Count;
                    DataRow dr = dt.Rows[i];

                    while (i < n && !dr["RowHeader2"].ToString().Equals(""))  // write first shift hour intervals
                    {

                        ws.Cells[r, c].Value = dr["RowHeader1"];
                        shift = dr["ShiftName"].ToString();
                        firstshift = dr["ShiftName"].ToString();

                        c++;
                        i++;
                        dr = dt.Rows[i];
                    }

                    while (shift.Equals(dr["ShiftName"].ToString()) && i < n)
                    {
                        i++;
                        dr = dt.Rows[i];
                    }

                    if (!firstshift.Equals(dr["ShiftName"].ToString())) // write second shift hour intervals
                    {
                        r = 7;
                        c = 2;
                        while (i < n && !dr["RowHeader2"].ToString().Equals(""))
                        {

                            ws.Cells[r, c].Value = dr["RowHeader1"];
                            shift = dr["ShiftName"].ToString();

                            c++;
                            i++;
                            dr = dt.Rows[i];
                        }

                        while (shift.Equals(dr["ShiftName"].ToString()) && i < n)
                        {
                            i++;
                            dr = dt.Rows[i];
                        }
                    }

                    if (!firstshift.Equals(dr["ShiftName"].ToString())) // write third shift hour intervals
                    {
                        r = 8;
                        c = 2;
                        while (i < n && !dr["RowHeader2"].ToString().Equals(""))
                        {
                            ws.Cells[r, c].Value = dr["RowHeader1"];
                            shift = dr["ShiftName"].ToString();

                            c++;
                            i++;
                            dr = dt.Rows[i];
                        }
                    }

                    r = 15;
                    c = 7;
                    i = 0;
                    dr = dt.Rows[i];

                    string rowheaderstart = dr["RowHeader3"].ToString();
                    string rowheaderend = dr["RowHeader3"].ToString();
                    int downtimecol = c;
                    do
                    {
                        ws.Cells[r, c].Value = dr["RowHeader3"];
                        if (dr["RowHeader3"].ToString().Equals("Down Time")) downtimecol = c;
                        c++;
                        i++;
                        dr = dt.Rows[i];
                    } while (i < n && !rowheaderstart.Equals(dr["RowHeader3"].ToString())); // Write headers: shift numbers, downtime, hourly tgt etc.

                    r = 17;
                    i = 0;
                    while (i < n) // fill from first column to last shift hour
                    {
                        c = 1;
                        dr = dt.Rows[i];
                        ws.Cells[r, c].Value = dr["Date"];
                        c++;
                        ws.Cells[r, c].Value = dr["ShiftName"];
                        c++;
                        ws.Cells[r, c].Value = dr["Machine"];
                        c++;
                        ws.Cells[r, c].Value = dr["Component"];
                        c++;
                        ws.Cells[r, c].Value = dr["Operation"];
                        c++;
                        ws.Cells[r, c].Value = dr["Operator"];
                        do
                        {
                            c++;
                            try
                            {
                                ws.Cells[r, c].Value = Convert.ToInt32(dr["rowValue"]);
                            }
                            catch
                            {
                                ws.Cells[r, c].Value = dr["rowValue"];
                            }
                            i++;
                            if (i >= n)
                                break;
                            dr = dt.Rows[i];
                        } while (!dr["Rowheader3"].ToString().Equals("Down Time"));

                        do
                        {
                            i++;
                            if (i >= n)
                                break;
                            dr = dt.Rows[i];
                        } while (!dr["Rowheader3"].ToString().Equals("Total output (%)"));
                        i++;
                        r++;
                    }

                    r = 17;
                    i = 0;

                    while (i < n) // fill in remaining details
                    {
                        dr = dt.Rows[i];
                        while (!dr["Rowheader3"].ToString().Equals("Down Time"))
                        {
                            i++;
                            dr = dt.Rows[i];
                        }
                        for (int j = 0; j < 5; j++)
                        {
                            try
                            {
                                ws.Cells[r, c + j + 1].Value = Convert.ToDouble(dt.Rows[i + j]["RowValue"]);
                            }
                            catch
                            {
                                ws.Cells[r, c + j + 1].Value = dt.Rows[i + j]["RowValue"];
                            }
                        }
                        i += 5;
                        r++;
                    }
                }
                // End sheet1, begin sheet2

                ws = excelPackage.Workbook.Worksheets[2];
                ws.Column(1).Style.Numberformat.Format = "dd-mmm-yy";
                string comparisonParam = "";
                string timeAxis = "Shift";
                string shiftName = "";
                string type = "Console";
                rdr = AccessReportData.GetProductionEfficiency(strtTime, endTime, plantid, MachineId, comparisonParam, timeAxis, shiftName, type);
                dt = new System.Data.DataTable();
                dt.Load(rdr);
                rdr.Close();
                r = 6;

                ws.Cells[3, 1].Value = strtTime;
                ws.Cells[3, 2].Value = endTime;
                ws.Cells[3, 3].Value = plantid.Equals("") ? "All" : plantid;
                ws.Cells[3, 4].Value = MachineId.Equals("") ? "All" : MachineId;
                ws.Cells[3, 5].Value = "All";

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        c = 1;
                        ws.Cells[r, c].Value = dr["Pdt"];
                        c++;
                        ws.Cells[r, c].Value = dr["Shftnm"];
                        c++;
                        ws.Cells[r, c].Value = dr["MachineID"];
                        c++;
                        ws.Cells[r, c].Value = dr["AE"];
                        c++;
                        ws.Cells[r, c].Value = dr["PE"];
                        c++;
                        ws.Cells[r, c].Value = dr["OE"];
                        r++;
                    }
                }
                // End sheet2, begin sheet3
                ws = excelPackage.Workbook.Worksheets[3];
                ws.Column(1).Style.Numberformat.Format = "dd-mmm-yy";
                string shiftID = "";
                string sheetNo = "3";
                string format = "1";
                rdr = AccessReportData.GetMandoReport(strtTime, endTime, plantid, MachineId, shiftID, sheetNo, format);
                dt = new System.Data.DataTable();
                dt.Load(rdr);
                rdr.Close();
                r = 7;

                ws.Cells[3, 1].Value = strtTime;
                ws.Cells[3, 2].Value = endTime;
                ws.Cells[3, 3].Value = plantid.Equals("") ? "All" : plantid;
                ws.Cells[3, 4].Value = MachineId.Equals("") ? "All" : MachineId;
                ws.Cells[3, 5].Value = "All";

                if (dt.Rows.Count > 0)
                {
                    //dataAvailable = true;
                    foreach (DataRow dr in dt.Rows)
                    {
                        c = 1;
                        ws.Cells[r, c].Value = dr["Date"];
                        c++;
                        ws.Cells[r, c].Value = dr["ShiftName"];
                        c++;
                        ws.Cells[r, c].Value = dr["Machine"];
                        c++;
                        ws.Cells[r, c].Value = dr["DownID"];
                        c++;
                        ws.Cells[r, c].Value = dr["DownTime"];
                        r++;
                    }
                }

                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("HourlyMachinewiseProductionReport Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile); //g: 
                }
                else
                {
                    Logger.WriteDebugLog("HourlyMachinewiseProductionReport not mailed: no production data");
                }
            }

            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.Message);
                Logger.WriteErrorLog(ex.StackTrace);
            }

            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        internal static void ExportToolLifeReport(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {
            //strtTime = "2018-02-26 06:00:00";//g: testdates
            //endTime = "2018-02-27 06:00:00";
            string dst = string.Empty;
            try
            {
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                dst = Path.Combine(ExportPath, string.Format("SM_ToolLifeReport_{0:ddMMMyyyy}.xlsx", DateTime.Parse(strtTime)));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);



                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                System.Data.DataTable dt = AccessReportData.GetToolLifeData(strtTime, endTime, MachineId);
                double minPercent = 50;
                if (dt.Rows.Count > 0)
                {
                    dataAvailable = true;
                    minPercent = AccessReportData.GetToolLifeThreshold();
                }

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Name = "Tool Life";

                ws.Cells["C3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm:ss");
                ws.Cells["F3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm:ss");

                int r = 7;
                foreach (DataRow row in dt.Rows)
                {
                    ws.Cells[r, 1].Value = row["Date"];
                    ws.Cells[r, 2].Value = row["Line"];
                    ws.Cells[r, 3].Value = row["Machine"];
                    ws.Cells[r, 4].Value = row["Tool"];
                    ws.Cells[r, 5].Value = row["ToolDescription"];
                    ws.Cells[r, 6].Value = row["ToolTarget"];
                    ws.Cells[r, 7].Value = row["ToolActual"];
                    ws.Cells[r, 8].Value = row["ChangeReason"];
                    ws.Cells[r, 9].Value = row["ToolPercent"];
                    if (Convert.ToDouble(row["ToolPercent"]) < minPercent && Convert.ToDouble(row["ToolPercent"]) >= 10)
                    {
                        ws.Cells[r, 9].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[r, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FF6347"));
                    }
                    r++;
                }
                //ws.Cells[6, 1].LoadFromDataTable(dt, true);

                // Update some of the header names
                //ws.Cells[6, 5].Value = "Target Tool Life";
                //ws.Cells[6, 6].Value = "Actual Tool Life";
                //ws.Cells[6, 7].Value = "Reason For Change";
                //ws.Cells[6, 8].Value = "% Tool Life Achieved";
                //ws.Row(1).Style.Font.Bold = true;
                //ws.Row(1).Style.Font.Size = 12;
                ws.Cells[7, 2, r, 8].AutoFitColumns();
                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("ToolLifeReport Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("ToolLifeReport not mailed: no data");
                }

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        internal static void ExportEWSOEEReport(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            string dst = string.Empty;
            //strtTime = "2018-02-26 06:00:00";//g: testdates
            //endTime = "2018-02-27 06:00:00";
            try
            {
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                if (isDayWise)
                {
                    dst = Path.Combine(ExportPath, string.Format("SM_EWSOEEReport_{0:ddMMMyyyy}.xlsx", DateTime.Parse(strtTime)));
                }
                else
                {
                    dst = Path.Combine(ExportPath, string.Format("SM_EWSOEEReport_{0:ddMMMyyyy_HHmmss}.xlsx", DateTime.Parse(strtTime)));
                }
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                System.Data.DataTable dt = AccessReportData.GetEWSOEEData(strtTime, endTime, MachineId, plantid);
                if (dt.Rows.Count > 0) dataAvailable = true;
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Name = "EWS OEE";

                ws.Cells["C3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm:ss");
                ws.Cells["F3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm:ss");

                int r = 7;
                string prevGrp = "";
                int prevRow = r;
                foreach (DataRow row in dt.Rows)
                {
                    ws.Cells[r, 1].Value = row["Date"];
                    ws.Cells[r, 2].Value = row["Line"];
                    ws.Cells[r, 3].Value = row["MachineID"];
                    if (!prevGrp.Equals(row["Group"].ToString()))
                    {
                        ws.Cells[r, 4].Value = row["Group"];
                        prevGrp = row["Group"].ToString();
                        if (prevRow != r)
                        {
                            ws.Cells[prevRow, 4, r - 1, 4].Merge = true;
                            ws.Cells[prevRow, 4, r - 1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        prevRow = r;
                    }
                    ws.Cells[r, 5].Value = row["target"];
                    ws.Cells[r, 6].Value = row["Actual"];
                    ws.Cells[r, 7].Value = row["OEE"];
                    ws.Cells[r, 8].Value = row["Reason"];
                    r++;
                }

                if (prevRow != r)
                {
                    ws.Cells[prevRow, 4, r - 1, 4].Merge = true;
                    ws.Cells[prevRow, 4, r - 1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                //ws.Cells[6, 1].LoadFromDataTable(dt, true);
                // Update some of the header names
                //ws.Cells[6, 4].Value = "Target Count";
                //ws.Cells[6, 5].Value = "Actual Count";
                //ws.Row(1).Style.Font.Bold = true;
                //ws.Row(1).Style.Font.Size = 12;
                ws.Cells[6, 2, r, 8].AutoFitColumns();
                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("EWSOEEReport Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile); //g: 
                }
                else
                {
                    Logger.WriteDebugLog("EWSOEEReport not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        // g:
        private static void FillTempl(ref ExcelWorksheet ws, ref int row, ref int col,
            ref Dictionary<string, int> catrow, ref Dictionary<string, int> subcatrow, ref Dictionary<string, List<string>> catAndSubCat, ref string strtTime, ref int lastrow)
        {
            int cnt = 0;
            catrow.Clear();
            subcatrow.Clear();

            foreach (string k in catAndSubCat.Keys)
            {
                List<string> subcats = catAndSubCat[k];
                ws.Cells[row + cnt, 1].Value = k;
                ws.Cells[row + cnt, 1].Style.Font.Bold = true;
                ws.Cells[row + cnt, 1, row + cnt, 2].Merge = true;
                ws.Cells[row + cnt, 1, row + cnt, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row + cnt, 1, row + cnt, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                catrow.Add(k, row + cnt);
                cnt += 1;
                foreach (var l in subcats)
                {
                    ws.Cells[row + cnt, 2].Value = l;
                    ws.Cells[row + cnt, 1].Value = cnt - catrow.Keys.Count + 1;
                    subcatrow.Add(l, row + cnt);
                    cnt += 1;
                }

                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Merge = true;
                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Style.Font.Bold = true;
                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Value = "Cycles";
                ws.Cells[row + cnt - subcats.Count, 4, row + cnt - 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            lastrow = row + cnt;

            var modelTable = ws.Cells[3, 1, lastrow - 1, 17];
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.BorderAround(ExcelBorderStyle.Medium);

            DateTime tmpdt = Convert.ToDateTime(strtTime);
            row = 5;
            col = 5;
            for (int i = 0; i < 12; i++)
            {
                ws.Cells[row, col + i].Value = tmpdt.ToString("MMM-yy");
                tmpdt = tmpdt.AddMonths(1);
            }
            ws.Cells["E6:Q" + lastrow.ToString()].Value = "";
        }


        internal static void ExportWeeklyEWSOEEReport(string strtTime, string endTime, string Shift, string strReportFile, string ExportPath, string ExportedReportFile,
           string MachineId, string operators, string sttime,
           string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, int ShiftID)
        {
            string dst = string.Empty;
            //strtTime = "2018-06-04 06:00:00";
            //endTime = "2018-06-09 06:00:00";
            try
            {
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SM_EWSWeeklyOEEReport_{0:ddMMMyyyy}.xlsx", DateTime.Parse(strtTime)));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Name = "Weekly OEE";

                ws.Cells["C3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm:ss");
                ws.Cells["F3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm:ss");

                System.Data.DataTable dt = AccessReportData.GetEWSWeeklyOEEData(strtTime, endTime, MachineId, plantid);
                if (dt.Rows.Count > 0) dataAvailable = true;

                int r = 7;
                string prevGrp = "";
                int prevRow = r;

                DateTime stDate = DateTime.Parse(strtTime);
                for (int i = 0; i < 6; i++)
                {
                    ws.Cells[r - 1, 4 + i].Value = stDate.AddDays(i).ToString("yyyy-MMM-dd");
                }

                Dictionary<string, int> rowList = new Dictionary<string, int>();
                Dictionary<string, float> gantryList = new Dictionary<string, float>();
                Dictionary<string, float> roboList = new Dictionary<string, float>();
                float tot = 0;

                foreach (DataRow row in dt.Rows)
                {

                    ws.Cells[r, 1].Value = row["Line"];
                    ws.Cells[r, 2].Value = row["MachineID"];
                    if (!prevGrp.Equals(row["Group"].ToString()))
                    {
                        ws.Cells[r, 3].Value = row["Group"];
                        if (prevRow != r)
                        {
                            ws.Cells[prevRow, 3, r - 1, 3].Merge = true;
                            ws.Cells[prevRow, 3, r - 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[prevRow, 11, r - 1, 11].Merge = true;
                            ws.Cells[prevRow, 11, r - 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[prevRow, 11].Formula = string.Format("AVERAGE(J{0},J{1})", prevRow, r - 1);
                            //var tmprng = ws.Cells[prevRow, 10, r - 1, 10].Value;
                            tot = 0; // get average and save
                            for (int j = prevRow; j < r; j++)
                            {
                                tot += float.Parse(ws.Cells[j, 10].Value.ToString());
                            }

                            if (prevGrp.IndexOf("gantry", StringComparison.OrdinalIgnoreCase) > -1)
                            {
                                gantryList.Add(prevGrp, tot / (r - prevRow));
                            }
                            if (prevGrp.IndexOf("robo", StringComparison.OrdinalIgnoreCase) > -1)
                            {
                                roboList.Add(prevGrp, tot / (r - prevRow));
                            }
                        }

                        prevGrp = row["Group"].ToString();
                        prevRow = r;
                    }
                    ws.Cells[r, 4].Value = row[3];
                    ws.Cells[r, 5].Value = row[4];
                    ws.Cells[r, 6].Value = row[5];
                    ws.Cells[r, 7].Value = row[6];
                    ws.Cells[r, 8].Value = row[7];
                    ws.Cells[r, 9].Value = row[8];
                    ws.Cells[r, 10].Value = float.Parse(row["OEE"].ToString());
                    r++;
                }

                //if (!prevGrp.Equals("")) rowList.Add(prevGrp, prevRow);

                if (prevRow != r)
                {
                    ws.Cells[prevRow, 3, r - 1, 3].Merge = true;
                    ws.Cells[prevRow, 3, r - 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[prevRow, 11, r - 1, 11].Merge = true;
                    ws.Cells[prevRow, 11, r - 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[prevRow, 11].Formula = string.Format("AVERAGE(J{0},J{1})", prevRow, r - 1);
                    tot = 0; // get average and save
                    for (int j = prevRow; j < r; j++)
                    {
                        tot += float.Parse(ws.Cells[j, 10].Value.ToString());
                    }

                    if (prevGrp.IndexOf("gantry", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        gantryList.Add(prevGrp, tot / (r - prevRow));
                    }
                    if (prevGrp.IndexOf("robo", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        roboList.Add(prevGrp, tot / (r - prevRow));
                    }
                }

                for (int j = 99; j < 99 + gantryList.Count; j++)
                {
                    ws.Cells[j, 99].Value = gantryList.ElementAt(j - 99).Key.Split(new[] { ' ' })[0];
                    ws.Cells[j, 100].Value = gantryList.ElementAt(j - 99).Value;
                }

                for (int j = 199; j < 199 + roboList.Count; j++)
                {
                    ws.Cells[j, 99].Value = roboList.ElementAt(j - 199).Key.Split(new[] { ' ' })[0];
                    ws.Cells[j, 100].Value = roboList.ElementAt(j - 199).Value;
                }

                ws.Cells[6, 1, r, 10].AutoFitColumns();

                //int pixeltop = 300;
                int pixelleft = 1025;

                var chart = (ExcelBarChart)ws.Drawings.AddChart(string.Format("barChart{0}", 0), eChartType.ColumnClustered);
                chart.SetSize(700, 300);
                chart.SetPosition(100, pixelleft);
                chart.Title.Text = "Gantry OEE";
                chart.Series.Add(ExcelRange.GetAddress(99, 100, 99 + gantryList.Count, 100), ExcelRange.GetAddress(99, 99, 99 + gantryList.Count, 99));

                chart = (ExcelBarChart)ws.Drawings.AddChart(string.Format("barChart{0}", 1), eChartType.ColumnClustered);
                chart.SetSize(700, 300);
                chart.SetPosition(450, pixelleft);
                chart.Title.Text = "Robo OEE";
                chart.Series.Add(ExcelRange.GetAddress(199, 100, 199 + roboList.Count, 100), ExcelRange.GetAddress(199, 99, 199 + roboList.Count, 99));

                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("EWS OEE Weekly Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("EWS OEE Weekly Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        internal static void ExportProductionAndDowntimesReport(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            string dst = string.Empty;
            //strtTime = "2019-11-13"; // g: test 
            //endTime = "2019-11-14"; // g: test 
            try
            {
                Logger.WriteDebugLog(string.Format("Start Time={0}, End Time={1}", strtTime, endTime));
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SM_DailyProductionandDowntimeDetails_{1}_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now, isDayWise ? "Day" : "Shift"));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Template Copied to Export Path");
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                ws.Cells["B3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm tt");
                ws.Cells["D3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm tt");

                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }

                System.Data.DataTable dtMachinelist;
                string parameter = "Summary";
                System.Data.DataTable dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                if (dt != null && dt.Rows.Count > 0) dataAvailable = true;
                Logger.WriteDebugLog("Values obtain for Parameter- Summary");

                int r = 7;
                Dictionary<string, int> dctrows = new Dictionary<string, int>();
                foreach (DataRow row in dtMachinelist.Rows)
                {
                    if (r == 7)
                    {
                        ws.Name = row["Machineid"].ToString();
                    }
                    else
                    {
                        excelPackage.Workbook.Worksheets.Add(row["Machineid"].ToString(), ws);

                    }

                    ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                    ws.Cells["F3"].Value = plantid.Equals("") ? "ALL" : plantid;
                    ws.Cells["H3"].Value = row["Machineid"].ToString();

                    dctrows.Add(row["Machineid"].ToString(), 12);
                    r++;
                }
                r = 7;
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                        //ws.Cells["F3"].Value = plantid.Equals("") ? "ALL" : plantid;
                        //ws.Cells["H3"].Value = row["Machineid"].ToString();
                        ws.Cells[r, 1].Value = row["Totaltime"];
                        ws.Cells[r, 2].Value = row["Runtime"];
                        ws.Cells[r, 3].Value = row["NetDowntime"];
                        ws.Cells[r, 4].Value = row["PDT"];
                        ws.Cells[r, 5].Value = row["RuntimeEffy"];
                        //ws.Cells[r, 5].Value = row["Managementloss"];
                        //ws.Cells[6, 1, r + 1, 6].AutoFitColumns();
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                        // ignore machine if not in summary
                    }
                }


                //parameter = "Efficiency";
                //dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter);
                //r = 10;
                //foreach (DataRow row in dt.Rows)
                //{
                //    try
                //    {
                //        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                //        ws.Cells[r, 2].Value = row["CycleCount"];
                //        ws.Cells[r, 4].Value = row["ProductionEfficiency"];
                //        ws.Cells[r, 6].Value = row["AvailabilityEfficiency"];
                //        ws.Cells[r, 8].Value = row["OverAllEfficiency"];
                //    }
                //    catch (Exception ex)
                //    {
                //        // ignore machine if not in summary
                //    }
                //}

                parameter = "COLevelDetails";
                dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                Logger.WriteDebugLog("Values obtain for Parameter- COLevelDetails");
                r = 12;
                string prevmach = "";

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (!prevmach.Equals(row["PMachineid"].ToString()))
                        {
                            if (!prevmach.Equals("")) dctrows[prevmach] = r;
                            prevmach = row["PMachineid"].ToString();
                            //if (r != 12) ws.Cells[12, 1, r, 6].AutoFitColumns();
                            ws = excelPackage.Workbook.Worksheets[prevmach];
                            r = 12;
                        }
                        ws.Cells[r, 1].Value = row["ComponentID"];
                        ws.Cells[r, 2].Value = row["OperationNo"];
                        //ws.Cells[r, 3].Value = row["ProdCount"];
                        //ws.Cells[r, 4].Value = row["StdCycleTime"];
                        //ws.Cells[r, 5].Value = row["AvgCycleTime"];
                        //ws.Cells[r, 6].Value = row["SpeedRation"];
                        //ws.Cells[r, 7].Value = row["StdLoadUnload"];
                        //ws.Cells[r, 8].Value = row["AvgLoadUnload"];
                        //ws.Cells[r, 9].Value = row["LoadRation"];
                        r++;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }
                if (!prevmach.Equals("")) dctrows[prevmach] = r;

                //if (ws != null) ws.Cells[12, 1, r, 6].AutoFitColumns();
                Dictionary<string, int> dctrowsPrev = new Dictionary<string, int>();
                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dctrowsPrev.Add(machine, dctrows[machine]);
                    using (ExcelRange range = ws.Cells[10, 1, dctrows[machine] - 1, 2])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                }

                parameter = "InProcessProdCycle";
                dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                Logger.WriteDebugLog("Values obtain for Parameter- InProcessProdCycle");
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Production Details: (Running Part)";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "ComponentID";
                        ws.Cells[dctrows[curMachine], 2].Value = "OperationNo";
                        //ws.Cells[dctrows[curMachine], 3].Value = "LoadUnloadTime";
                        //ws.Cells[dctrows[curMachine], 4].Value = "Net CycleTime";
                        //ws.Cells[dctrows[curMachine], 5].Value = "ICD";
                        //ws.Cells[dctrows[curMachine], 6].Value = "PDT";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 6].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                prevmach = "";
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (!prevmach.Equals(row["Machineid"].ToString()))
                        {
                            prevmach = row["Machineid"].ToString();
                            ws = excelPackage.Workbook.Worksheets[prevmach];
                            //dctrows[prevmach] = r;
                        }
                        ws.Cells[dctrows[prevmach], 1].Value = row["ComponentID"];
                        ws.Cells[dctrows[prevmach], 2].Value = row["OperationNo"];
                        //ws.Cells[dctrows[prevmach], 3].Value = row["LoadUnloadTime"];
                        //ws.Cells[dctrows[prevmach], 4].Value = row["CycleTime"];
                        //ws.Cells[dctrows[prevmach], 5].Value = row["In_Cycle_DownTime"];
                        //ws.Cells[dctrows[prevmach], 6].Value = row["PDT"];
                        dctrows[prevmach] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    
                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 2])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    dctrowsPrev[machine] = dctrows[machine];
                }

                #region "DownTime Summery"

                parameter = "DowntimeSummary";
                //'2019-10-01 06:00:00','2019-10-02 06:00:00'
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Downtime Summary:";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "DownDescription";
                        ws.Cells[dctrows[curMachine], 2].Value = "Sum Downtime in min";
                        ws.Cells[dctrows[curMachine], 3].Value = "No of Occurences";
                        ws.Cells[dctrows[curMachine], 4].Value = "MIN. Downtime in min";
                        ws.Cells[dctrows[curMachine], 5].Value = "MAX. Downtime in min";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 5].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, machine, plantid, parameter, out dtMachinelist);
                    Logger.WriteDebugLog("Values obtain for Parameter- DowntimeSummary");
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            ws.Cells[dctrows[machine], 1].Value = row["DownDescription"];
                            ws.Cells[dctrows[machine], 2].Value = row["DownTime"];
                            ws.Cells[dctrows[machine], 3].Value = row["NoOfOccurences"];
                            ws.Cells[dctrows[machine], 4].Value = row["MinDowntime"];
                            ws.Cells[dctrows[machine], 5].Value = row["MaxDowntime"];
                            dctrows[machine] += 1;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }
                }
                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];

                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    dctrowsPrev[machine] = dctrows[machine];
                }
                #endregion

                parameter = "InProcessDownCycles";
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Downtime Details:";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "StartTime";
                        ws.Cells[dctrows[curMachine], 2].Value = "EndTime";
                        ws.Cells[dctrows[curMachine], 3].Value = "DownDescription ";
                        ws.Cells[dctrows[curMachine], 4].Value = "Net DownTime";
                        ws.Cells[dctrows[curMachine], 5].Value = "Down Status";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 5].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, machine, plantid, parameter, out dtMachinelist);
                    Logger.WriteDebugLog("Values obtain for Parameter- InProcessDownCycles");
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            ws.Cells[dctrows[machine], 1].Value = Convert.ToDateTime(row["StartTime"]).ToString("dd-MM-yy HH:mm:ss"); // Avoid showing in number format for date
                            ws.Cells[dctrows[machine], 2].Value = Convert.ToDateTime(row["EndTime"]).ToString("dd-MM-yy HH:mm:ss");
                            ws.Cells[dctrows[machine], 3].Value = row["DownDescription"];
                            ws.Cells[dctrows[machine], 4].Value = row["netDownTime"];
                            ws.Cells[dctrows[machine], 5].Value = row["DownStatus"];
                            dctrows[machine] += 1;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }

                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[3, 1, 3, 8])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[5, 1, 7, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[10, 1, dctrows[machine], 5])
                    {
                        range.AutoFitColumns();
                        ws.Cells["D3"].AutoFitColumns(); // autofit date

                    }
                    using (ExcelRange range = ws.Cells[1, 1, 1, 9])
                    {
                        range.Value = "DOWN TIME REPORT- " + (isDayWise ? "Day" : "Shiftwise");
                    }
                    using (ExcelRange range = ws.Cells[3, 4, 18, 5])
                    {
                        range.AutoFitColumns();
                    }
                }

                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("DailyProductionandDowntime Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("DailyProductionandDowntime Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        internal static void ExportProductionAndDowntimesReportWeeklyorMonthly(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isWeekly, out bool isexported)
        {
            string dst = string.Empty;
            isexported = false;
            //strtTime = "2019-11-13"; // g: test 
            //endTime = "2019-11-14"; // g: test 
            try
            {
                Logger.WriteDebugLog(string.Format("Start Time={0}, End Time={1}", strtTime, endTime));
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SM_DailyProductionandDowntimeDetails_{1}_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now, isWeekly ? "Week" : "Month"));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Template Copied to Export Path");
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                ws.Cells["B3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm tt");
                ws.Cells["D3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm tt");

                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }
                System.Data.DataTable dtMachinelist;
                string parameter = "Summary";
                System.Data.DataTable dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                if (dt != null && dt.Rows.Count > 0) dataAvailable = true;
                Logger.WriteDebugLog("Values obtain for Parameter- Summary");
                int r = 7;
                Dictionary<string, int> dctrows = new Dictionary<string, int>();
                foreach (DataRow row in dt.Rows)
                {
                    if (r == 7)
                    {
                        ws.Name = row["Machineid"].ToString();
                    }
                    else
                    {
                        excelPackage.Workbook.Worksheets.Add(row["Machineid"].ToString(), ws);
                    }
                    dctrows.Add(row["Machineid"].ToString(), 12);
                    r++;
                }

                r = 7;
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                        ws.Cells["F3"].Value = plantid.Equals("") ? "ALL" : plantid;
                        ws.Cells["H3"].Value = row["Machineid"].ToString();
                        ws.Cells[r, 1].Value = row["Totaltime"];
                        ws.Cells[r, 2].Value = row["Runtime"];
                        ws.Cells[r, 3].Value = row["NetDowntime"];
                        ws.Cells[r, 4].Value = row["PDT"];
                        ws.Cells[r, 5].Value = row["RuntimeEffy"];
                        //ws.Cells[r, 5].Value = row["Managementloss"];
                        //ws.Cells[6, 1, r + 1, 6].AutoFitColumns();
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                        // ignore machine if not in summary
                    }
                }


                //parameter = "Efficiency";
                //dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter);
                //r = 10;
                //foreach (DataRow row in dt.Rows)
                //{
                //    try
                //    {
                //        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                //        ws.Cells[r, 2].Value = row["CycleCount"];
                //        ws.Cells[r, 4].Value = row["ProductionEfficiency"];
                //        ws.Cells[r, 6].Value = row["AvailabilityEfficiency"];
                //        ws.Cells[r, 8].Value = row["OverAllEfficiency"];
                //    }
                //    catch (Exception ex)
                //    {
                //        // ignore machine if not in summary
                //    }
                //}

                parameter = "COLevelDetails";
                dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                Logger.WriteDebugLog("Values obtain for Parameter- COLevelDetails");
                r = 12;
                string prevmach = "";

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (!prevmach.Equals(row["PMachineid"].ToString()))
                        {
                            if (!prevmach.Equals("")) dctrows[prevmach] = r;
                            prevmach = row["PMachineid"].ToString();
                            //if (r != 12) ws.Cells[12, 1, r, 6].AutoFitColumns();
                            ws = excelPackage.Workbook.Worksheets[prevmach];
                            r = 12;
                        }
                        ws.Cells[r, 1].Value = row["ComponentID"];
                        ws.Cells[r, 2].Value = row["OperationNo"];
                        //ws.Cells[r, 3].Value = row["ProdCount"];
                        //ws.Cells[r, 4].Value = row["StdCycleTime"];
                        //ws.Cells[r, 5].Value = row["AvgCycleTime"];
                        //ws.Cells[r, 6].Value = row["SpeedRation"];
                        //ws.Cells[r, 7].Value = row["StdLoadUnload"];
                        //ws.Cells[r, 8].Value = row["AvgLoadUnload"];
                        //ws.Cells[r, 9].Value = row["LoadRation"];
                        r++;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }
                if (!prevmach.Equals("")) dctrows[prevmach] = r;

                //if (ws != null) ws.Cells[12, 1, r, 6].AutoFitColumns();
                Dictionary<string, int> dctrowsPrev = new Dictionary<string, int>();
                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dctrowsPrev.Add(machine, dctrows[machine]);
                    using (ExcelRange range = ws.Cells[10, 1, dctrows[machine] - 1, 2])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                }

                parameter = "InProcessProdCycle";
                dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                Logger.WriteDebugLog("Values obtain for Parameter- InProcessProdCycle");
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Production Details: (Running Part)";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "ComponentID";
                        ws.Cells[dctrows[curMachine], 2].Value = "OperationNo";
                        //ws.Cells[dctrows[curMachine], 3].Value = "LoadUnloadTime";
                        //ws.Cells[dctrows[curMachine], 4].Value = "Net CycleTime";
                        //ws.Cells[dctrows[curMachine], 5].Value = "ICD";
                        //ws.Cells[dctrows[curMachine], 6].Value = "PDT";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 6].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                prevmach = "";
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (!prevmach.Equals(row["Machineid"].ToString()))
                        {
                            prevmach = row["Machineid"].ToString();
                            ws = excelPackage.Workbook.Worksheets[prevmach];
                            //dctrows[prevmach] = r;
                        }
                        ws.Cells[dctrows[prevmach], 1].Value = row["ComponentID"];
                        ws.Cells[dctrows[prevmach], 2].Value = row["OperationNo"];
                        //ws.Cells[dctrows[prevmach], 3].Value = row["LoadUnloadTime"];
                        //ws.Cells[dctrows[prevmach], 4].Value = row["CycleTime"];
                        //ws.Cells[dctrows[prevmach], 5].Value = row["In_Cycle_DownTime"];
                        //ws.Cells[dctrows[prevmach], 6].Value = row["PDT"];
                        dctrows[prevmach] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];

                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 2])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    dctrowsPrev[machine] = dctrows[machine];
                }

                #region "DownTime Summery"

                parameter = "DowntimeSummary";
                //'2019-10-01 06:00:00','2019-10-02 06:00:00'
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Downtime Summary:";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "DownDescription";
                        ws.Cells[dctrows[curMachine], 2].Value = "Sum Downtime in min";
                        ws.Cells[dctrows[curMachine], 3].Value = "No of Occurences";
                        ws.Cells[dctrows[curMachine], 4].Value = "MIN. Downtime in min";
                        ws.Cells[dctrows[curMachine], 5].Value = "MAX. Downtime in min";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 5].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                    Logger.WriteDebugLog("Values obtain for Parameter- DowntimeSummary");
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            ws.Cells[dctrows[machine], 1].Value = row["DownDescription"];
                            ws.Cells[dctrows[machine], 2].Value = row["DownTime"];
                            ws.Cells[dctrows[machine], 3].Value = row["NoOfOccurences"];
                            ws.Cells[dctrows[machine], 4].Value = row["MinDowntime"];
                            ws.Cells[dctrows[machine], 5].Value = row["MaxDowntime"];
                            dctrows[machine] += 1;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }
                }
                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];

                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    dctrowsPrev[machine] = dctrows[machine];
                }
                #endregion

                parameter = "InProcessDownCycles";
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Downtime Details:";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "StartTime";
                        ws.Cells[dctrows[curMachine], 2].Value = "EndTime";
                        ws.Cells[dctrows[curMachine], 3].Value = "DownDescription ";
                        ws.Cells[dctrows[curMachine], 4].Value = "Net DownTime";
                        ws.Cells[dctrows[curMachine], 5].Value = "Down Status";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 5].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                    Logger.WriteDebugLog("Values obtain for Parameter- InProcessDownCycles");
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            ws.Cells[dctrows[machine], 1].Value = Convert.ToDateTime(row["StartTime"]).ToString("dd-MM-yy HH:mm:ss"); // Avoid showing in number format for date
                            ws.Cells[dctrows[machine], 2].Value = Convert.ToDateTime(row["EndTime"]).ToString("dd-MM-yy HH:mm:ss");
                            ws.Cells[dctrows[machine], 3].Value = row["DownDescription"];
                            ws.Cells[dctrows[machine], 4].Value = row["netDownTime"];
                            ws.Cells[dctrows[machine], 5].Value = row["DownStatus"];
                            dctrows[machine] += 1;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }

                    using (ExcelRange range = ws.Cells[dctrowsPrev[machine] + 1, 1, dctrows[machine] - 1, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[3, 1, 3, 8])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[5, 1, 7, 5])
                    {
                        //range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }

                    using (ExcelRange range = ws.Cells[10, 1, dctrows[machine], 5])
                    {
                        range.AutoFitColumns();
                        ws.Cells["D3"].AutoFitColumns(); // autofit date

                    }
                    using (ExcelRange range = ws.Cells[1, 1, 1, 9])
                    {
                        range.Value = "DOWN TIME REPORT- " + (isWeekly ? "Weekly" : "Monthly");
                    }
                    using (ExcelRange range = ws.Cells[3, 4, 18, 5])
                    {
                        range.AutoFitColumns();
                    }
                }

                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    isexported = true;
                    //Logger.WriteDebugLog("DailyProductionandDowntime Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("DailyProductionandDowntime Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        internal static void ExportOEEAndLosstimeReport(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            string dst = string.Empty;
            bool dataAvailable = false;
            try
            {
                strtTime = DateTime.Parse(strtTime).ToString("yyyy-MM-dd");
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                Directory.CreateDirectory(ExportPath);
                dst = Path.Combine(ExportPath, string.Format("OEEAndLosstime_{1}_{0}.xlsx", strtTime, MachineId));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                ws.Name = "OEE And Losstime";

                List<System.Data.DataTable> dtables = AccessReportData.GetOEEAndLosstime(MachineId, strtTime, endTime);
                ws.Cells["AC1"].Value = string.Format("DATE: {0}", strtTime); //.ToString("dd-MMM-yyyy"));

                int rownum = 4;
                int colnum = 4;
                int rownumBac = rownum;
                System.Data.DataTable dt = dtables[0];

                for (int j = 0; j < 19; j++)
                {
                    if (colnum == 12 || colnum == 18)
                        colnum++;
                    ws.Cells[rownum, colnum].Value = dt.Rows[j]["Downid"].ToString().ToUpper();
                    colnum++;
                }

                dt = dtables[1];
                rownum = 5;
                foreach (DataRow row in dt.Rows)
                {
                    ws.Cells[rownum, 1].Value = row["MachineID"];
                    ws.Cells[rownum, 2].Value = row["AvlTotalTime"];
                    ws.Cells[rownum, 3].Value = row["AvlTime"];
                    ws.Cells[rownum, 4].Value = Convert.ToDouble(row["D1"]);
                    ws.Cells[rownum, 5].Value = Convert.ToDouble(row["D2"]);
                    ws.Cells[rownum, 6].Value = Convert.ToDouble(row["D3"]);
                    ws.Cells[rownum, 7].Value = Convert.ToDouble(row["D4"]);
                    ws.Cells[rownum, 8].Value = Convert.ToDouble(row["D5"]);
                    ws.Cells[rownum, 9].Value = Convert.ToDouble(row["D6"]);
                    ws.Cells[rownum, 10].Value = Convert.ToDouble(row["D7"]);
                    ws.Cells[rownum, 11].Value = Convert.ToDouble(row["D8"]);
                    ws.Cells[rownum, 12].Value = Convert.ToDouble(row["LoadingTime"]);
                    ws.Cells[rownum, 13].Value = Convert.ToDouble(row["D9"]);
                    ws.Cells[rownum, 14].Value = Convert.ToDouble(row["D10"]);
                    ws.Cells[rownum, 15].Value = Convert.ToDouble(row["D11"]);
                    ws.Cells[rownum, 16].Value = Convert.ToDouble(row["D12"]);
                    ws.Cells[rownum, 17].Value = Convert.ToDouble(row["D13"]);
                    ws.Cells[rownum, 18].Value = Convert.ToDouble(row["OperatingTime"]);
                    ws.Cells[rownum, 19].Value = Convert.ToDouble(row["D14"]);
                    ws.Cells[rownum, 20].Value = Convert.ToDouble(row["D15"]);
                    ws.Cells[rownum, 21].Value = Convert.ToDouble(row["D16"]);
                    ws.Cells[rownum, 22].Value = Convert.ToDouble(row["D17"]);
                    ws.Cells[rownum, 23].Value = Convert.ToDouble(row["D18"]);
                    ws.Cells[rownum, 24].Value = Convert.ToDouble(row["D19"]);
                    ws.Cells[rownum, 25].Value = Convert.ToDouble(row["NetOperatingTime"]);
                    ws.Cells[rownum, 26].Value = row["Hold"];
                    ws.Cells[rownum, 27].Value = row["RejMat"];
                    ws.Cells[rownum, 28].Value = row["RejPro"];
                    ws.Cells[rownum, 29].Value = row["ValuableOperatingTime"];
                    ws.Cells[rownum, 30].Value = row["AEffy"];
                    ws.Cells[rownum, 31].Value = row["PEffy"];
                    ws.Cells[rownum, 32].Value = row["QEffy"];
                    ws.Cells[rownum, 33].Value = row["OEffy"];
                    rownum++;
                }

                if (rownum > 5)
                {
                    dataAvailable = true;
                }

                var rng = ws.Cells[5, 1, rownum, 33];
                rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rng.Style.Font.Size = 9;
                rng.Style.Font.Bold = true;

                ws.Cells[5, 1, rownum, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 2, rownum, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 4, rownum, 12].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 13, rownum, 18].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 19, rownum, 25].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 26, rownum, 29].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[5, 30, rownum, 33].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                ws.Cells[rownum, 1, rownum, 33].Style.Border.BorderAround(ExcelBorderStyle.Medium);



                dt = dtables[2];
                foreach (DataRow row in dt.Rows)
                {
                    ws.Cells[rownum, 1].Value = "Total/Avgt";
                    ws.Cells[rownum, 2].Value = row["Tot_AvlTotalTime"];
                    ws.Cells[rownum, 3].Value = row["Tot_AvlTime"];
                    ws.Cells[rownum, 4].Value = Convert.ToDouble(row["D1"]);
                    ws.Cells[rownum, 5].Value = Convert.ToDouble(row["D2"]);
                    ws.Cells[rownum, 6].Value = Convert.ToDouble(row["D3"]);
                    ws.Cells[rownum, 7].Value = Convert.ToDouble(row["D4"]);
                    ws.Cells[rownum, 8].Value = Convert.ToDouble(row["D5"]);
                    ws.Cells[rownum, 9].Value = Convert.ToDouble(row["D6"]);
                    ws.Cells[rownum, 10].Value = Convert.ToDouble(row["D7"]);
                    ws.Cells[rownum, 11].Value = Convert.ToDouble(row["D8"]);
                    ws.Cells[rownum, 12].Value = Convert.ToDouble(row["LoadingTime"]);
                    ws.Cells[rownum, 13].Value = Convert.ToDouble(row["D9"]);
                    ws.Cells[rownum, 14].Value = Convert.ToDouble(row["D10"]);
                    ws.Cells[rownum, 15].Value = Convert.ToDouble(row["D11"]);
                    ws.Cells[rownum, 16].Value = Convert.ToDouble(row["D12"]);
                    ws.Cells[rownum, 17].Value = Convert.ToDouble(row["D13"]);
                    ws.Cells[rownum, 18].Value = Convert.ToDouble(row["OperatingTime"]);
                    ws.Cells[rownum, 19].Value = Convert.ToDouble(row["D14"]);
                    ws.Cells[rownum, 20].Value = Convert.ToDouble(row["D15"]);
                    ws.Cells[rownum, 21].Value = Convert.ToDouble(row["D16"]);
                    ws.Cells[rownum, 22].Value = Convert.ToDouble(row["D17"]);
                    ws.Cells[rownum, 23].Value = Convert.ToDouble(row["D18"]);
                    ws.Cells[rownum, 24].Value = Convert.ToDouble(row["D19"]);
                    ws.Cells[rownum, 25].Value = Convert.ToDouble(row["NetOperatingTime"]);
                    ws.Cells[rownum, 26].Value = row["Tot_Hold"];
                    ws.Cells[rownum, 27].Value = row["Tot_RejMat"];
                    ws.Cells[rownum, 28].Value = row["Tot_RejPro"];
                    ws.Cells[rownum, 29].Value = row["ValuableOperatingTime"];
                    ws.Cells[rownum, 30].Value = row["Tot_AEffy"];
                    ws.Cells[rownum, 31].Value = row["Tot_PEffy"];
                    ws.Cells[rownum, 32].Value = row["Tot_QEffy"];
                    ws.Cells[rownum, 33].Value = row["Tot_OEffy"];
                    rownum++;
                }
                rownumBac = rownum;
                dt = dtables[3];
                rownum = 17;

                foreach (DataRow row in dt.Rows)
                {
                    ws.Cells[rownum, 1].Value = string.Format("Available Time: {0:0.##}% of total time", row["AvailableTime"]).ToUpper();
                    ws.Cells[rownum, 30].Value = string.Format("Plant Closure: {0:0.##}% of total time", row["PlantClosureTime"]).ToUpper();
                    ws.Cells[rownum + 1, 1].Value = string.Format("Loading Time: {0:0.##}% of total time", row["LoadingTime"]).ToUpper();
                    ws.Cells[rownum + 1, 19].Value = string.Format("Others (P, A, M, RM): {0:0.##}% of total time", row["Others"]).ToUpper();
                    ws.Cells[rownum + 1, 25].Value = string.Format("No Pron Planned: {0:0.##}% of total time", row["NoPronPlanned"]).ToUpper();
                    ws.Cells[rownum + 2, 1].Value = string.Format("Operating Time: {0:0.##}% of loading time", row["OperatingTime"]).ToUpper();
                    ws.Cells[rownum + 2, 14].Value = string.Format("Downtime: {0:0.##}% of Avl. Time", row["DownTime"]).ToUpper();
                    ws.Cells[rownum + 3, 1].Value = string.Format("Net opt. time: {0:0.##}% of loading Time", row["NetOperatingTime"]).ToUpper();
                    ws.Cells[rownum + 3, 10].Value = string.Format("speed loss: {0:0.##}% of Avl. Time", row["SpeedLoss"]).ToUpper();
                    ws.Cells[rownum + 4, 1].Value = string.Format("Val. time: {0:0.##}% of loading Time", row["ValuableOperatingTime"]).ToUpper();
                    ws.Cells[rownum + 4, 6].Value = string.Format("quality loss: {0:0.##}% of Avl. Time", row["QualityLoss"]).ToUpper();
                    rownum++;
                }

                //ws.DeleteRow(rownumBac, 15 - rownumBac, true);
                for (int i = rownumBac; i < 15; i++)
                {
                    ws.Row(i).Height = 0;
                }
                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("OEEAndLosstime Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("OEEAndLosstime Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
            }
        }

        internal static void ExportMangalDowntimeReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC)
        {
            string dst = string.Empty;
            bool dataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("MangalDowntime_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }


                File.Copy(strReportFile, dst, true);

                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
                worksheet.Name = "Mangal Downtime";

                System.Data.DataTable toprows = new System.Data.DataTable();
                System.Data.DataTable bottomrows = new System.Data.DataTable();
                System.Data.DataTable data = new System.Data.DataTable();
                System.Data.DataTable lastcolumn = new System.Data.DataTable();
                data = AccessReportData.GetmangalDowntime(sttime, ndtime, out toprows, out bottomrows, out lastcolumn);
                if (data.Rows.Count > 0)
                    dataAvailable = true;
                worksheet.Cells["B4"].Value = sttime.ToString("dd-MM-yyyy");
                worksheet.Cells["F4"].Value = ndtime.ToString("dd-MM-yyyy");
                int row = 10, col = 1;
                for (int i = 0; i < toprows.Rows.Count; i++)
                {
                    worksheet.Cells[row, col].Value = toprows.Rows[i]["Downid"];
                    row++;
                }
                row = 25;
                for (int i = 0; i < bottomrows.Rows.Count; i++)
                {
                    worksheet.Cells[row, col].Value = bottomrows.Rows[i]["downCatergory"];
                    row++;
                }
                row = 7; col = 3;
                foreach (DataRow item in data.Rows)
                {
                    row = 7;
                    worksheet.Cells[row, col].Value = item["MachineID"];
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["AvailableTime"].ToString());
                    row = row + 2;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D1"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D2"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D3"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D4"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D5"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["NetAvalTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["Utilization"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["RunTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["Uptime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["AvailabilityEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["ProductionEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["QualityEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["OverallEfficiency"].ToString());
                    row = row + 2;
                    worksheet.Cells[row, col].Value = getdoubledata(item["TotalDelayTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C1"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C2"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C3"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C4"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C5"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C6"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C7"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C8"].ToString());
                    col++;
                }

                worksheet.Cells[7, col].Value = "All machines" + "\n" + " Up time";
                row = 8;
                foreach (DataRow item in lastcolumn.Rows)
                {
                    worksheet.Cells[row, col].Value = getdoubledata(item["AvailableTime"].ToString());
                    row = row + 2;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D1"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D2"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D3"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D4"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["D5"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["NetAvalTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["Utilization"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["RunTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["Uptime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["AvailabilityEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["ProductionEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["QualityEfficiency"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["OverallEfficiency"].ToString());
                    row = row + 2;
                    worksheet.Cells[row, col].Value = getdoubledata(item["TotalDelayTime"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C1"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C2"].ToString()); ;
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C3"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C4"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C5"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C6"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C7"].ToString());
                    row++;
                    worksheet.Cells[row, col].Value = getdoubledata(item["C8"].ToString());
                }
                worksheet.Cells[6, 1, row, col + 1].AutoFitColumns();
                worksheet.Cells[16, 1, 16, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[16, 1, 16, col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                worksheet.Cells[18, 1, 18, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[18, 1, 18, col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                worksheet.Cells[7, 1, row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[7, 1, row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[7, 1, row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[7, 1, row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[7, 1, row, col].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[7, 1, row, col].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[7, 1, row, col].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[7, 1, row, col].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);

                ExcelPieChart pieChart = worksheet.Drawings.AddChart("Forging production analysis Pie Chart", eChartType.Pie3D) as ExcelPieChart;
                pieChart.Title.Text = "Forging production analysis Pie Chart";
                pieChart.Series.Add(ExcelRange.GetAddress(25, col, 32, col), ExcelRange.GetAddress(25, 1, 32, 1));
                pieChart.Legend.Position = eLegendPosition.Bottom;
                pieChart.DataLabel.ShowValue = true;
                pieChart.SetSize(600, 600);
                pieChart.SetPosition(40, 0, 2, 0);

                excelPackage.SaveAs(newFile);
                if (dataAvailable)
                {
                    Logger.WriteDebugLog("MangalDowntime Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("MangalDowntime Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);

            }
        }
        private static double getdoubledata(string datastring)
        {
            double value = 0.00;
            Double.TryParse(datastring, out value);
            return value;
        }


        internal static void ExportEfficiencyAndGraphReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string machineid)
        {
            //sttime = Convert.ToDateTime("2019-05-02");
            //ndtime = Convert.ToDateTime("2019-05-03");
            //machineid = "CNC-01";
            string dst = string.Empty;
            const int INITCOL = 3;
            bool dataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("BaluAutoEfficiencyReport_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                System.Data.DataTable dtDownlist = new System.Data.DataTable();
                System.Data.DataTable dtEff = new System.Data.DataTable();
                System.Data.DataTable dtTotal = new System.Data.DataTable();
                var copyworksheet1 = excelPackage.Workbook.Worksheets[1];
                var copyworksheet2 = excelPackage.Workbook.Worksheets[2];
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                #region "For All Machine"
                if (machineid.Equals("ALL", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(machineid))
                {
                    List<string> mid = new List<string>();
                    mid = AccessReportData.GetMachineid();
                    string MID = mid[0];
                    foreach (string id in mid)
                    {
                        AccessReportData.GetEfficiencyAndGraphReport(id, sttime, ndtime, out dtDownlist, out dtEff, out dtTotal);
                        //AccessReportData.GetEfficiencyAndGraphReportMonthly(sttime, out dtDownlist, out dtEff, out dtTotal);
                        if (MID != id)
                        {
                            ws = excelPackage.Workbook.Worksheets.Add(id + " Graph", copyworksheet1);
                        }
                        ws.Name = id + " Graph";
                        int colnum = INITCOL;
                        int rownum = 3;
                        double num = 0;
                        string curDate = "";
                        int prevcol = 3;
                        int curcol = 3;
                        ExcelRange rng;
                        bool firstIter = true;
                        string prevDate = sttime.ToString("dd-MMM");
                        Dictionary<string, double> dctDowns = new Dictionary<string, double>();
                        if (dtTotal.Rows.Count > 0)
                        {
                            if (!dataAvailable)
                                dataAvailable = true;
                            DataRow tmp = dtTotal.Rows[0];
                            foreach (DataRow row in dtDownlist.Rows)
                            {
                                if (double.TryParse(tmp[string.Format("D{0}", rownum - 2)].ToString(), out num))
                                {
                                    dctDowns.Add(string.Format("{0} - {1}", rownum - 2, row["Downid"].ToString()), num);
                                }
                                else
                                {
                                    Logger.WriteDebugLog("Some downtimes are not in decimal format");
                                    break;
                                }
                                rownum += 1;
                            }
                            if (double.TryParse(tmp["Others"].ToString(), out num))
                            {
                                dctDowns.Add(string.Format("{0} - Other Losses", rownum - 2), num);
                            }
                            else
                            {
                                Logger.WriteDebugLog("Some downtimes are not in decimal format");
                            }
                            rownum = 3;
                            foreach (KeyValuePair<string, double> item in dctDowns.OrderByDescending(key => key.Value))
                            {
                                ws.Cells[rownum, 1].Value = item.Key;
                                ws.Cells[rownum, 2].Value = item.Value;
                                rownum += 1;
                            }

                            ws.Cells[rownum, 1].Value = "Total Losses";
                            ws.Cells[rownum, 2].Value = double.TryParse(tmp["TotalLoss"].ToString(), out num) ? num : tmp["TotalLoss"];
                        }

                        var barchartAE = ws.Drawings["Chart 1"] as ExcelBarChart;
                        barchartAE.Series.Delete(0);
                        barchartAE.Series.Add(ExcelRange.GetAddress(3, 2, 17, 2), ExcelRange.GetAddress(3, 1, 17, 1));

                        rownum = 10;
                        if (MID != id)
                        {
                            ws = excelPackage.Workbook.Worksheets.Add(id, copyworksheet2);
                        }
                        else
                        {
                            ws = excelPackage.Workbook.Worksheets[2];
                        }
                        ws.Name = id;
                        ws.Cells["A1"].Value = string.Format("OEE Details: {1} ({0}-{2})", sttime.ToString("dd-MMM-yyyy"), id, ndtime.ToString("dd-MMM-yyyy"));
                        foreach (DataRow row in dtDownlist.Rows)
                        {
                            ws.Cells[rownum, 1].Value = string.Format("{0} - {1}", rownum - 9, row[1].ToString());
                            rownum += 1;
                        }

                        #region "Plotting Efficiency values machinewise"
                        foreach (DataRow row in dtEff.Rows)
                        {
                            curDate = ((DateTime)row["Day"]).ToString("dd-MMM");
                            if (!prevDate.Equals(curDate) && firstIter)
                            {
                                prevDate = curDate;
                            }
                            firstIter = false;
                            if (prevDate != curDate)
                            {
                                rng = ws.Cells[2, prevcol, 2, colnum - 1];
                                rng.Merge = true;
                                rng.Value = prevDate;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                prevDate = curDate;
                                prevcol = colnum;
                            }
                            ws.Cells[2, colnum].Value = curDate;
                            ws.Cells[3, colnum].Value = row["Shift"].ToString();
                            ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                            ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                            ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                            ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                            ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                            ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                            ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                            ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                            ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                            ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                            ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                            ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                            ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                            ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                            ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                            ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                            ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                            ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                            ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                            ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                            ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                            ws.Cells[28, colnum].Value = row["PartCycleTime"];
                            ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                            ws.Cells[30, colnum].Value = row["UtilisedTime"];
                            ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                            ws.Cells[32, colnum].Value = row["QuantityProduced"];
                            ws.Cells[33, colnum].Value = row["QuantityRejected"];
                            ws.Cells[34, colnum].Value = row["QuantityOK"];
                            colnum += 1;
                        }
                        if (colnum - prevcol > 1)
                        {
                            rng = ws.Cells[2, prevcol, 2, colnum - 1];
                            rng.Merge = true;
                            rng.Value = prevDate;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            prevDate = curDate;
                            prevcol = colnum;
                        }
                        #endregion

                        #region "Plotting Total Values"
                        foreach (DataRow row in dtTotal.Rows)
                        {
                            ws.Cells[2, colnum].Value = "Total";
                            ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                            ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                            ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                            ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                            ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                            ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                            ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                            ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                            ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                            ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                            ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                            ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                            ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                            ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                            ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                            ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                            ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                            ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                            ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                            ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                            ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                            ws.Cells[28, colnum].Value = row["PartCycleTime"];
                            ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                            ws.Cells[30, colnum].Value = row["UtilisedTime"];
                            ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                            ws.Cells[32, colnum].Value = row["QuantityProduced"];
                            ws.Cells[33, colnum].Value = row["QuantityRejected"];
                            ws.Cells[34, colnum].Value = row["QuantityOK"];
                            colnum += 1;
                            //firstIter = false;
                        }
                        #endregion

                        #region "Assigning Colors to excel cells"
                        if (colnum > INITCOL)
                        {
                            // formatting the table
                            // 
                            rng = ws.Cells[2, 1, 34, colnum - 1];
                            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            rng.AutoFitColumns();

                            // first two columns
                            rng = ws.Cells[2, 1, 34, 2];
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                            // remaining columns
                            rng = ws.Cells[2, 3, 2, colnum - 1];
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                            // green
                            rownum = 4;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                            // green
                            rownum = 9;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                            // green
                            rownum = 27;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                            rownum = 34;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));

                            // shades of gray
                            rownum = 24;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                            rownum = 25;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                            rownum = 30;
                            rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xcc, 0xcc, 0xcc));

                            // light green
                            rng = ws.Cells[5, 3, 8, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x5e, 0xf2, 0x68));

                            // total column: yellow
                            rng = ws.Cells[5, colnum - 1, 34, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xfc, 0xfc, 0x25));
                            rng.Style.Font.Bold = true;

                            // headings
                            rng = ws.Cells[2, 3, 3, colnum - 1];
                            rng.Style.Font.Bold = true;

                            // clay
                            rng = ws.Cells[26, 1, 26, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xef, 0xd5, 0x9b));

                            // blue
                            rng = ws.Cells[31, 1, 31, colnum - 1];
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x47, 0x57, 0xe8));
                            rng.Style.Font.Color.SetColor(Color.White);

                            // font size 10
                            rng = ws.Cells[4, 1, 34, colnum - 1];
                            rng.Style.Font.Size = 10;
                        }
                        //ws.Column(5).Hidden = true;
                        //ws.Column(8).Hidden = true;
                        #endregion
                        MID = id;
                    }
                }
                #endregion

                #region "For Single Machine"
                else
                {
                    AccessReportData.GetEfficiencyAndGraphReport(machineid, sttime, ndtime, out dtDownlist, out dtEff, out dtTotal);
                    //AccessReportData.GetEfficiencyAndGraphReportMonthly(sttime, out dtDownlist, out dtEff, out dtTotal);
                    ws.Name = machineid + " Graph";
                    int colnum = INITCOL;
                    int rownum = 3;
                    double num = 0;
                    string curDate = "";
                    int prevcol = 3;
                    int curcol = 3;
                    ExcelRange rng;
                    bool firstIter = true;
                    string prevDate = sttime.ToString("dd-MMM");
                    Dictionary<string, double> dctDowns = new Dictionary<string, double>();
                    if (dtTotal.Rows.Count > 0)
                    {
                        if (!dataAvailable)
                            dataAvailable = true;
                        DataRow tmp = dtTotal.Rows[0];
                        foreach (DataRow row in dtDownlist.Rows)
                        {
                            if (double.TryParse(tmp[string.Format("D{0}", rownum - 2)].ToString(), out num))
                            {
                                dctDowns.Add(string.Format("{0} - {1}", rownum - 2, row["Downid"].ToString()), num);
                            }
                            else
                            {
                                Logger.WriteDebugLog("Some downtimes are not in decimal format");
                                break;
                            }
                            rownum += 1;
                        }
                        if (double.TryParse(tmp["Others"].ToString(), out num))
                        {
                            dctDowns.Add(string.Format("{0} - Other Losses", rownum - 2), num);
                        }
                        else
                        {
                            Logger.WriteDebugLog("Some downtimes are not in decimal format");
                        }
                        rownum = 3;
                        foreach (KeyValuePair<string, double> item in dctDowns.OrderByDescending(key => key.Value))
                        {
                            ws.Cells[rownum, 1].Value = item.Key;
                            ws.Cells[rownum, 2].Value = item.Value;
                            rownum += 1;
                        }

                        ws.Cells[rownum, 1].Value = "Total Losses";
                        ws.Cells[rownum, 2].Value = double.TryParse(tmp["TotalLoss"].ToString(), out num) ? num : tmp["TotalLoss"];
                    }

                    var barchartAE = ws.Drawings["Chart 1"] as ExcelBarChart;
                    barchartAE.Series.Delete(0);
                    barchartAE.Series.Add(ExcelRange.GetAddress(3, 2, 17, 2), ExcelRange.GetAddress(3, 1, 17, 1));

                    rownum = 10;
                    ws = excelPackage.Workbook.Worksheets[2];
                    ws.Name = machineid;
                    ws.Cells["A1"].Value = string.Format("OEE Details: {1} ({0}-{2})", sttime.ToString("dd-MMM-yyyy"), machineid, ndtime.ToString("dd-MMM-yyyy"));
                    foreach (DataRow row in dtDownlist.Rows)
                    {
                        ws.Cells[rownum, 1].Value = string.Format("{0} - {1}", rownum - 9, row[1].ToString());
                        rownum += 1;
                    }

                    #region "Plotting Efficiency values"
                    foreach (DataRow row in dtEff.Rows)
                    {
                        curDate = ((DateTime)row["Day"]).ToString("dd-MMM");
                        if (!prevDate.Equals(curDate) && firstIter)
                        {
                            prevDate = curDate;
                        }
                        firstIter = false;
                        if (prevDate != curDate)
                        {
                            rng = ws.Cells[2, prevcol, 2, colnum - 1];
                            rng.Merge = true;
                            rng.Value = prevDate;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            prevDate = curDate;
                            prevcol = colnum;
                        }
                        ws.Cells[2, colnum].Value = curDate;
                        ws.Cells[3, colnum].Value = row["Shift"].ToString();
                        ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                        ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                        ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                        ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                        ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                        ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                        ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                        ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                        ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                        ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                        ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                        ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                        ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                        ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                        ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                        ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                        ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                        ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                        ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                        ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                        ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                        ws.Cells[28, colnum].Value = row["PartCycleTime"];
                        ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                        ws.Cells[30, colnum].Value = row["UtilisedTime"];
                        ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                        ws.Cells[32, colnum].Value = row["QuantityProduced"];
                        ws.Cells[33, colnum].Value = row["QuantityRejected"];
                        ws.Cells[34, colnum].Value = row["QuantityOK"];
                        colnum += 1;
                    }
                    if (colnum - prevcol > 1)
                    {
                        rng = ws.Cells[2, prevcol, 2, colnum - 1];
                        rng.Merge = true;
                        rng.Value = prevDate;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        prevDate = curDate;
                        prevcol = colnum;
                    }
                    #endregion

                    #region "Plotting Total Values"
                    foreach (DataRow row in dtTotal.Rows)
                    {
                        ws.Cells[2, colnum].Value = "Total";
                        ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                        ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                        ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                        ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                        ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                        ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                        ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                        ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                        ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                        ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                        ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                        ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                        ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                        ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                        ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                        ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                        ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                        ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                        ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                        ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                        ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                        ws.Cells[28, colnum].Value = row["PartCycleTime"];
                        ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                        ws.Cells[30, colnum].Value = row["UtilisedTime"];
                        ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                        ws.Cells[32, colnum].Value = row["QuantityProduced"];
                        ws.Cells[33, colnum].Value = row["QuantityRejected"];
                        ws.Cells[34, colnum].Value = row["QuantityOK"];
                        colnum += 1;
                        //firstIter = false;
                    }
                    #endregion

                    if (colnum > INITCOL)
                    {
                        // formatting the table
                        // 
                        rng = ws.Cells[2, 1, 34, colnum - 1];
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        rng.AutoFitColumns();


                        // first two columns
                        rng = ws.Cells[2, 1, 34, 2];
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        // remaining columns
                        rng = ws.Cells[2, 3, 2, colnum - 1];
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        // green
                        rownum = 4;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                        // green
                        rownum = 9;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                        // green
                        rownum = 27;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                        rownum = 34;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));

                        // shades of gray
                        rownum = 24;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                        rownum = 25;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                        rownum = 30;
                        rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xcc, 0xcc, 0xcc));

                        // light green
                        rng = ws.Cells[5, 3, 8, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x5e, 0xf2, 0x68));

                        // total column: yellow
                        rng = ws.Cells[5, colnum - 1, 34, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xfc, 0xfc, 0x25));
                        rng.Style.Font.Bold = true;

                        // headings
                        rng = ws.Cells[2, 3, 3, colnum - 1];
                        rng.Style.Font.Bold = true;

                        // clay
                        rng = ws.Cells[26, 1, 26, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xef, 0xd5, 0x9b));

                        // blue
                        rng = ws.Cells[31, 1, 31, colnum - 1];
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x47, 0x57, 0xe8));
                        rng.Style.Font.Color.SetColor(Color.White);

                        // font size 10
                        rng = ws.Cells[4, 1, 34, colnum - 1];
                        rng.Style.Font.Size = 10;
                    }
                    //ws.Column(5).Hidden = true;
                    //ws.Column(8).Hidden = true;
                }
                #endregion

                if (dataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("EfficiencyAndGraph Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("EfficiencyAndGraph Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
                throw;
            }
        }

        internal static void ExportEfficiencyAndGraphReportMonthwise(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string machineid)
        {
            //sttime = Convert.ToDateTime("2020-02-15");
            string dst = string.Empty;
            const int INITCOL = 3;
            bool dataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("BaluAutoEfficiencyReport_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                //File.Copy(strReportFile, dst, true);
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                System.Data.DataTable dtDownlist = new System.Data.DataTable();
                System.Data.DataTable dtEff = new System.Data.DataTable();
                System.Data.DataTable dtTotal = new System.Data.DataTable();
                System.Data.DataTable dtAvgTotal = new System.Data.DataTable();
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
                AccessReportData.GetEfficiencyAndGraphReportMonthly(sttime, out dtDownlist, out dtEff, out dtTotal, out dtAvgTotal);
                ws.Name = "Graph";
                int colnum = INITCOL;
                int rownum = 3;
                double num = 0;
                string curDate = "";
                int prevcol = 3;
                //int curcol = 3;
                ExcelRange rng;
                bool firstIter = true;
                string prevDate = sttime.ToString("dd-MMM");
                Dictionary<string, double> dctDowns = new Dictionary<string, double>();

                if (dtTotal.Rows.Count > 0)
                {
                    if (!dataAvailable)
                        dataAvailable = true;
                    DataRow tmp = dtTotal.Rows[0];
                    foreach (DataRow row in dtDownlist.Rows)
                    {
                        if (double.TryParse(tmp[string.Format("D{0}", rownum - 2)].ToString(), out num))
                        {
                            dctDowns.Add(string.Format("{0} - {1}", rownum - 2, row["Downid"].ToString()), num);
                        }
                        else
                        {
                            Logger.WriteDebugLog("Some downtimes are not in decimal format");
                            break;
                        }
                        rownum += 1;
                    }
                    if (double.TryParse(tmp["Others"].ToString(), out num))
                    {
                        dctDowns.Add(string.Format("{0} - Other Losses", rownum - 2), num);
                    }
                    else
                    {
                        Logger.WriteDebugLog("Some downtimes are not in decimal format");
                    }
                    rownum = 3;
                    foreach (KeyValuePair<string, double> item in dctDowns.OrderByDescending(key => key.Value))
                    {
                        ws.Cells[rownum, 1].Value = item.Key;
                        ws.Cells[rownum, 2].Value = item.Value;
                        rownum += 1;
                    }

                    ws.Cells[rownum, 1].Value = "Total Losses";
                    ws.Cells[rownum, 2].Value = double.TryParse(tmp["TotalLoss"].ToString(), out num) ? num : tmp["TotalLoss"];
                }

                var barchartAE = ws.Drawings["Chart 1"] as ExcelBarChart;
                barchartAE.Series.Delete(0);
                barchartAE.Series.Add(ExcelRange.GetAddress(3, 2, 17, 2), ExcelRange.GetAddress(3, 1, 17, 1));

                rownum = 10;
                ws = excelPackage.Workbook.Worksheets[2];
                ws.Name = "MONTH";
                ws.Cells["A1"].Value = string.Format("OEE Details: {0} ", sttime.ToString("MMM"));
                foreach (DataRow row in dtDownlist.Rows)
                {
                    ws.Cells[rownum, 1].Value = string.Format("{0} - {1}", rownum - 9, row[1].ToString());
                    rownum += 1;
                }

                #region "Plotting Efficiency values machinewise"
                foreach (DataRow row in dtEff.Rows)
                {
                    curDate = ((DateTime)row["Day"]).ToString("dd-MMM");
                    if (!prevDate.Equals(curDate) && firstIter)
                    {
                        prevDate = curDate;
                    }
                    firstIter = false;
                    if (prevDate != curDate)
                    {
                        rng = ws.Cells[2, prevcol, 2, colnum - 1];
                        rng.Merge = true;
                        rng.Value = prevDate;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        prevDate = curDate;
                        prevcol = colnum;
                    }
                    ws.Cells[2, colnum].Value = curDate;
                    ws.Cells[3, colnum].Value = row["Shift"].ToString();
                    ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                    ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                    ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                    ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                    ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                    ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                    ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                    ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                    ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                    ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                    ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                    ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                    ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                    ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                    ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                    ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                    ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                    ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                    ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                    ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                    ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                    ws.Cells[28, colnum].Value = row["PartCycleTime"];
                    ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                    ws.Cells[30, colnum].Value = row["UtilisedTime"];
                    ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                    ws.Cells[32, colnum].Value = row["QuantityProduced"];
                    ws.Cells[33, colnum].Value = row["QuantityRejected"];
                    ws.Cells[34, colnum].Value = row["QuantityOK"];
                    colnum += 1;
                }
                if (colnum - prevcol > 1)
                {
                    rng = ws.Cells[2, prevcol, 2, colnum - 1];
                    rng.Merge = true;
                    rng.Value = prevDate;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    prevDate = curDate;
                    prevcol = colnum;
                }
                #endregion

                #region"Plotting Avg Total Values"
                foreach (DataRow row in dtAvgTotal.Rows)
                {
                    ws.Cells[2, colnum].Value = " Average Total";
                    ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                    ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                    ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                    ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                    ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                    ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                    ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                    ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                    ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                    ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                    ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                    ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                    ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                    ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                    ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                    ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                    ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                    ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                    ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                    ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                    ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                    ws.Cells[28, colnum].Value = row["PartCycleTime"];
                    ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                    ws.Cells[30, colnum].Value = row["UtilisedTime"];
                    ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                    ws.Cells[32, colnum].Value = row["QuantityProduced"];
                    ws.Cells[33, colnum].Value = row["QuantityRejected"];
                    ws.Cells[34, colnum].Value = row["QuantityOK"];
                    colnum += 1;
                    //firstIter = false;
                }
                #endregion

                #region "Plotting Total Values"
                foreach (DataRow row in dtTotal.Rows)
                {
                    ws.Cells[2, colnum].Value = "Total";
                    ws.Cells[5, colnum].Value = row["OverallEfficiency"];
                    ws.Cells[6, colnum].Value = row["AvailabilityEfficiency"];
                    ws.Cells[7, colnum].Value = row["ProductionEfficiency"];
                    ws.Cells[8, colnum].Value = row["QualityEfficiency"];
                    ws.Cells[10, colnum].Value = double.TryParse(row["D1"].ToString(), out num) ? num : row["D1"];
                    ws.Cells[11, colnum].Value = double.TryParse(row["D2"].ToString(), out num) ? num : row["D2"];
                    ws.Cells[12, colnum].Value = double.TryParse(row["D3"].ToString(), out num) ? num : row["D3"];
                    ws.Cells[13, colnum].Value = double.TryParse(row["D4"].ToString(), out num) ? num : row["D4"];
                    ws.Cells[14, colnum].Value = double.TryParse(row["D5"].ToString(), out num) ? num : row["D5"];
                    ws.Cells[15, colnum].Value = double.TryParse(row["D6"].ToString(), out num) ? num : row["D6"];
                    ws.Cells[16, colnum].Value = double.TryParse(row["D7"].ToString(), out num) ? num : row["D7"];
                    ws.Cells[17, colnum].Value = double.TryParse(row["D8"].ToString(), out num) ? num : row["D8"];
                    ws.Cells[18, colnum].Value = double.TryParse(row["D9"].ToString(), out num) ? num : row["D9"];
                    ws.Cells[19, colnum].Value = double.TryParse(row["D10"].ToString(), out num) ? num : row["D10"];
                    ws.Cells[20, colnum].Value = double.TryParse(row["D11"].ToString(), out num) ? num : row["D11"];
                    ws.Cells[21, colnum].Value = double.TryParse(row["D12"].ToString(), out num) ? num : row["D12"];
                    ws.Cells[22, colnum].Value = double.TryParse(row["D13"].ToString(), out num) ? num : row["D13"];
                    ws.Cells[23, colnum].Value = double.TryParse(row["Others"].ToString(), out num) ? num : row["Others"];
                    ws.Cells[24, colnum].Value = double.TryParse(row["TotalLoss"].ToString(), out num) ? num : row["TotalLoss"];
                    ws.Cells[25, colnum].Value = double.TryParse(row["ActLoss"].ToString(), out num) ? num : row["ActLoss"];
                    ws.Cells[26, colnum].Value = double.TryParse(row["LossErr"].ToString(), out num) ? num : row["LossErr"];
                    ws.Cells[28, colnum].Value = row["PartCycleTime"];
                    ws.Cells[29, colnum].Value = row["PlannedProductionTime"];
                    ws.Cells[30, colnum].Value = row["UtilisedTime"];
                    ws.Cells[31, colnum].Value = row["QuantityPlanned"];
                    ws.Cells[32, colnum].Value = row["QuantityProduced"];
                    ws.Cells[33, colnum].Value = row["QuantityRejected"];
                    ws.Cells[34, colnum].Value = row["QuantityOK"];
                    colnum += 1;
                    //firstIter = false;
                }
                #endregion

                #region "Assigning Colors to excel cells"
                if (colnum > INITCOL)
                {
                    // formatting the table
                    // 
                    rng = ws.Cells[2, 1, 34, colnum - 1];
                    rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    rng.AutoFitColumns();

                    // first two columns
                    rng = ws.Cells[2, 1, 34, 2];
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // remaining columns
                    rng = ws.Cells[2, 3, 2, colnum - 1];
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // green
                    rownum = 4;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                    // green
                    rownum = 9;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                    // green
                    rownum = 27;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));
                    rownum = 34;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x3a, 0x99, 0x40));

                    // shades of gray
                    rownum = 24;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                    rownum = 25;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xee, 0xee, 0xee));
                    rownum = 30;
                    rng = ws.Cells[rownum, 1, rownum, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xcc, 0xcc, 0xcc));

                    // light green
                    rng = ws.Cells[5, 3, 8, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x5e, 0xf2, 0x68));

                    // total column: yellow
                    rng = ws.Cells[5, colnum - 1, 34, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xfc, 0xfc, 0x25));
                    rng.Style.Font.Bold = true;

                    // headings
                    rng = ws.Cells[2, 3, 3, colnum - 1];
                    rng.Style.Font.Bold = true;

                    // clay
                    rng = ws.Cells[26, 1, 26, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xef, 0xd5, 0x9b));

                    // blue
                    rng = ws.Cells[31, 1, 31, colnum - 1];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x47, 0x57, 0xe8));
                    rng.Style.Font.Color.SetColor(Color.White);

                    // font size 10
                    rng = ws.Cells[4, 1, 34, colnum - 1];
                    rng.Style.Font.Size = 10;
                }
                //ws.Column(5).Hidden = true;
                //ws.Column(8).Hidden = true;
                #endregion

                if (dataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("EfficiencyAndGraph Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("EfficiencyAndGraph Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
            }
        }

        internal static void ExportMangalHourlychartReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string machineid, string plantId)
        {
            string dst = string.Empty;
            bool dataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("MangalHourlychart_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }

                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet worksheet = null;
                #region "For All Machine"
                if (machineid.Equals("ALL", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(machineid))
                {
                    List<string> mid = new List<string>();
                    mid = AccessReportData.GetMachineid();
                    //var copyworksheet = excelPackage.Workbook.Worksheets[1];
                    //worksheet = excelPackage.Workbook.Worksheets[1];
                    //worksheet.Name = "MangalHourlychart " + mid[0];
                    int Count = 1;
                    foreach (string id in mid)
                    {
                        worksheet = excelPackage.Workbook.Worksheets[Count++];
                        int shift = 1;
                        string shiftname = "", HourID = "";
                        System.Data.DataTable dt = new System.Data.DataTable();
                        System.Data.DataTable dtloss = new System.Data.DataTable();

                        worksheet.Cells["AK11"].Value = id;
                        worksheet.Cells["AP11"].Value = sttime.ToShortDateString();
                        dt = AccessReportData.ShiftProductionCountHourlyBNG(sttime, plantId, id);
                        if (dt.Rows.Count > 0)
                            dataAvailable = true;
                        int row = 16, col = 2;
                        foreach (DataRow item in dt.Rows)
                        {
                            col = 2;
                            int.TryParse(item["ShiftID"].ToString(), out shift);
                            switch (shift)
                            {
                                case 1:
                                    worksheet.Cells[row, col].Value = item["Target"];
                                    col = col + 2;
                                    worksheet.Cells[row, col].Value = item["Actual"];
                                    row++;
                                    break;
                                case 2:
                                    row++;
                                    worksheet.Cells[row, col].Value = item["Target"];
                                    col = col + 2;
                                    worksheet.Cells[row, col].Value = item["Actual"];
                                    break;
                                case 3:
                                    if (shiftname != item["ShiftName"].ToString())
                                        row = row + 2;
                                    else
                                        row++;
                                    worksheet.Cells[row, col].Value = item["Target"];
                                    col = col + 2;
                                    worksheet.Cells[row, col].Value = item["Actual"];
                                    break;
                            }
                            shiftname = item["ShiftName"].ToString();
                        }
                        dtloss = AccessReportData.ShiftProductionCountHourlyBNGAeeLoss(sttime, plantId, id);
                        if (dtloss != null && dtloss.Rows.Count > 0)
                        {
                            row = 15; col = 36;
                            for (int i = 0; i <= 8; i++)
                            {
                                worksheet.Cells[row, col].Value = dtloss.Rows[i]["DownCategory"].ToString();
                                col++;
                            }
                            row = 15; col = 36;
                            foreach (DataRow item in dtloss.Rows)
                            {

                                int.TryParse(item["ShiftID"].ToString(), out shift);
                                switch (shift)
                                {
                                    case 1:
                                        if (HourID != item["HourID"].ToString())
                                        {
                                            row++;
                                            col = 36;
                                            HourID = item["HourID"].ToString();
                                        }

                                        worksheet.Cells[row, col].Value = item["DownTime"];
                                        col++;
                                        break;
                                    case 2:
                                        if (item["HourID"].ToString() == "0")
                                            row = row + 1;
                                        else
                                        {
                                            if (HourID != item["HourID"].ToString())
                                            {
                                                row++;
                                                col = 36;
                                                HourID = item["HourID"].ToString();
                                            }
                                            worksheet.Cells[row, col].Value = item["DownTime"];
                                            col++;
                                        }
                                        break;
                                    case 3:
                                        if (item["HourID"].ToString() == "0")
                                        {
                                            row++;
                                        }
                                        else
                                        {
                                            if (HourID != item["HourID"].ToString())
                                            {
                                                row++;
                                                col = 36;
                                                HourID = item["HourID"].ToString();
                                            }
                                            worksheet.Cells[row, col].Value = item["DownTime"];
                                            col++;
                                        }
                                        break;
                                }
                                shiftname = item["ShiftID"].ToString();
                            }

                            worksheet.Cells[15, 36, 15, col - 1].AutoFitColumns();
                            //worksheet.Workbook.FullCalcOnLoad = true;
                            worksheet.Workbook.CalcMode = ExcelCalcMode.Automatic;
                            worksheet.Calculate();
                        }
                        //worksheet.Name = "HourlyChart" + id.Trim().Replace(" ", "");
                    }
                    for (int x = mid.Count + 1; x <= 60; x++)
                    {
                        try
                        {
                            excelPackage.Workbook.Worksheets.Delete(mid.Count + 1);
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.Message);

                        }
                    }
                }
                #endregion
                #region "For Single Machine"
                else
                {
                    worksheet = excelPackage.Workbook.Worksheets[1];
                    int shift = 1;
                    string shiftname = "", HourID = "";
                    System.Data.DataTable dt = new System.Data.DataTable();
                    System.Data.DataTable dtloss = new System.Data.DataTable();

                    worksheet.Cells["AK11"].Value = machineid;
                    worksheet.Cells["AP11"].Value = sttime.ToShortDateString();
                    dt = AccessReportData.ShiftProductionCountHourlyBNG(sttime, plantId, machineid);
                    if (dt.Rows.Count > 0)
                        dataAvailable = true;
                    int row = 16, col = 2;
                    foreach (DataRow item in dt.Rows)
                    {
                        col = 2;
                        int.TryParse(item["ShiftID"].ToString(), out shift);
                        switch (shift)
                        {
                            case 1:
                                worksheet.Cells[row, col].Value = item["Target"];
                                col = col + 2;
                                worksheet.Cells[row, col].Value = item["Actual"];
                                row++;
                                break;
                            case 2:
                                row++;
                                worksheet.Cells[row, col].Value = item["Target"];
                                col = col + 2;
                                worksheet.Cells[row, col].Value = item["Actual"];
                                break;
                            case 3:
                                if (shiftname != item["ShiftName"].ToString())
                                    row = row + 2;
                                else
                                    row++;
                                worksheet.Cells[row, col].Value = item["Target"];
                                col = col + 2;
                                worksheet.Cells[row, col].Value = item["Actual"];
                                break;
                        }
                        shiftname = item["ShiftName"].ToString();
                    }
                    dtloss = AccessReportData.ShiftProductionCountHourlyBNGAeeLoss(sttime, plantId, machineid);
                    if (dtloss != null && dtloss.Rows.Count > 0)
                    {
                        row = 15; col = 36;
                        for (int i = 0; i <= 8; i++)
                        {
                            worksheet.Cells[row, col].Value = dtloss.Rows[i]["DownCategory"].ToString();
                            col++;
                        }
                        row = 15; col = 36;
                        foreach (DataRow item in dtloss.Rows)
                        {

                            int.TryParse(item["ShiftID"].ToString(), out shift);
                            switch (shift)
                            {
                                case 1:
                                    if (HourID != item["HourID"].ToString())
                                    {
                                        row++;
                                        col = 36;
                                        HourID = item["HourID"].ToString();
                                    }

                                    worksheet.Cells[row, col].Value = item["DownTime"];
                                    col++;
                                    break;
                                case 2:
                                    if (item["HourID"].ToString() == "0")
                                        row = row + 1;
                                    else
                                    {
                                        if (HourID != item["HourID"].ToString())
                                        {
                                            row++;
                                            col = 36;
                                            HourID = item["HourID"].ToString();
                                        }
                                        worksheet.Cells[row, col].Value = item["DownTime"];
                                        col++;
                                    }
                                    break;
                                case 3:
                                    if (item["HourID"].ToString() == "0")
                                    {
                                        row++;
                                    }
                                    else
                                    {
                                        if (HourID != item["HourID"].ToString())
                                        {
                                            row++;
                                            col = 36;
                                            HourID = item["HourID"].ToString();
                                        }
                                        worksheet.Cells[row, col].Value = item["DownTime"];
                                        col++;
                                    }
                                    break;
                            }
                            shiftname = item["ShiftID"].ToString();

                        }

                        worksheet.Cells[15, 36, 15, col - 1].AutoFitColumns();
                        //worksheet.Workbook.FullCalcOnLoad = true;
                        worksheet.Workbook.CalcMode = ExcelCalcMode.Automatic;
                        worksheet.Calculate();
                    }
                    for (int x = 2; x <= 60; x++)
                    {
                        try
                        {
                            excelPackage.Workbook.Worksheets.Delete(2);
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.Message);
                        }
                    }
                }
                #endregion
                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("MangalHourlychart Report Exported successfully");

                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, newFile.ToString(), ExportedReportFile);
                }

                else
                {
                    Logger.WriteDebugLog("MangalHourlychart Report not mailed: no data");
                }

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
        }

        internal static void ExportSonaMISReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string PlantID, string Shift)
        {
            //sttime = Convert.ToDateTime("2019-09-12");
            //ndtime = Convert.ToDateTime("2019-09-13");
            string dst = string.Empty;
            bool status = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SonaMISReport_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                //FileInfo newFile = new FileInfo(dst);

                System.Data.DataTable downTable = new System.Data.DataTable();
                System.Data.DataTable shiftwiseData = new System.Data.DataTable();
                List<string> AllShift = new List<string>();

                AccessReportData.GetSONAMISReportData(sttime, ndtime, PlantID, Shift, out downTable, out shiftwiseData);
                AllShift = AccessReportData.GetAllShift();
                if (downTable != null && downTable.Rows.Count > 0 && shiftwiseData != null && shiftwiseData.Rows.Count > 0)
                {
                    status = Reports.MISReport(dst, AllShift, PlantID, shiftwiseData, downTable, sttime, ndtime);
                }

                if (status)
                {
                    Logger.WriteDebugLog("Sona MIS Report Exported successfully");

                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }

                else
                {
                    Logger.WriteDebugLog("Sona MIS Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
        }

        internal static void ExportFlowMeterReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC, string PlantID, string MachineID, string Shift)
        {
            //sttime = Convert.ToDateTime("2019-12-04 00:00:00.000");
            //ndtime = Convert.ToDateTime("2019-12-05 00:00:00.000");
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("FlowMeterReport_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                System.Data.DataTable dt = new System.Data.DataTable();
                dt = AccessReportData.GetFlowMeterReportData(sttime, ndtime, PlantID, MachineID);
                if (dt != null && dt.Rows.Count > 0)
                {
                    Logger.WriteDebugLog("FlowMeter Report generation has started...............");
                    DataAvailable = true;
                    int ROW = 8;
                    foreach (DataRow row in dt.Rows)
                    {
                        ws.Cells[ROW, 1].Value = row["StartTime"];
                        ws.Cells[ROW, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[ROW, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        ws.Cells[ROW, 2].Value = row["Endtime"];
                        ws.Cells[ROW, 2].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        ws.Cells[ROW, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        if (string.IsNullOrEmpty(PlantID) || PlantID.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                        {
                            ws.Cells[ROW, 3].Value = row["PlantID"];
                            ws.Cells[ROW, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            ws.Cells[ROW, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        if (string.IsNullOrEmpty(MachineID) || MachineID.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                        {
                            ws.Cells[ROW, 4].Value = row["machineid"];
                            ws.Cells[ROW, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            ws.Cells[ROW, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        if (string.IsNullOrEmpty(Shift) || Shift.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                        {
                            ws.Cells[ROW, 5].Value = row["Shift"];
                            ws.Cells[ROW, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            ws.Cells[ROW, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        ws.Cells[ROW, 6].Value = row["componentid"];
                        ws.Cells[ROW, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[ROW, 7].Value = row["operationno"];
                        ws.Cells[ROW, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[ROW, 8].Value = row["Flowvalue1"];
                        ws.Cells[ROW, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[ROW, 9].Value = row["Flowvalue2"];
                        ws.Cells[ROW, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[ROW, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ROW++;
                    }
                    ws.Cells["B3"].Value = sttime;
                    ws.Cells["B3"].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                    if (!(string.IsNullOrEmpty(PlantID) || PlantID.Equals("ALL", StringComparison.OrdinalIgnoreCase)))
                    {
                        ws.DeleteColumn(3);
                        ws.Cells["B5"].Value = PlantID;
                    }
                    else
                        ws.Cells["B5"].Value = "All";

                    if (!(string.IsNullOrEmpty(MachineID) || MachineID.Equals("ALL", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (ws.Cells["C7"].Value.ToString() == "Plant ID")
                            ws.DeleteColumn(4);
                        else if (ws.Cells["C7"].Value.ToString() == "Machine ID")
                            ws.DeleteColumn(3);

                        ws.Cells["D5"].Value = MachineID;
                    }
                    else
                        ws.Cells["D5"].Value = "All";

                    if (string.IsNullOrEmpty(Shift) || Shift.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                    {
                        ws.Cells["F5"].Value = "All";
                    }
                    else
                    {
                        if (ws.Cells["C7"].Value.ToString() == "Plant ID" && ws.Cells["D7"].Value.ToString() == "Machine ID")
                        {
                            ws.DeleteColumn(5);
                            ws.Cells["B5"].Value = "All";
                            ws.Cells["D5"].Value = "All";
                        }
                        else if (ws.Cells["C7"].Value.ToString() == "Plant ID" && ws.Cells["D7"].Value.ToString() == "Shift")
                        {
                            ws.DeleteColumn(4);
                            ws.Cells["B5"].Value = "All";
                            ws.Cells["D5"].Value = MachineID;
                        }
                        else if (ws.Cells["C7"].Value.ToString() == "Machine ID" && ws.Cells["D7"].Value.ToString() == "Shift")
                        {
                            ws.DeleteColumn(4);
                            ws.Cells["D5"].Value = "All";
                            ws.Cells["B5"].Value = PlantID;
                        }
                        else if (ws.Cells["C7"].Value.ToString() == "Shift")
                        {
                            ws.DeleteColumn(3);
                            ws.Cells["D5"].Value = MachineID;
                        }

                        ws.Cells["F5"].Value = Shift;
                    }
                    ws.Cells["C3"].Value = "To:";
                    ws.Cells["C3"].Style.Font.Bold = true;
                    ws.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["C5"].Value = "MachineID:";
                    ws.Cells["C5"].Style.Font.Bold = true;
                    ws.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["D3"].Value = ndtime;
                    ws.Cells["D3"].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                    ws.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["E5"].Value = "Shift:";
                    ws.Cells["E5"].Style.Font.Bold = true;
                    ws.Cells["E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells["B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells["F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }

                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Flow Meter Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("Flow Meter Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
        }
        internal static void ExportEnergyMeterReport(string strReportFile, string ExportPath, string ExportedReportFile,
        int ExportType, DateTime sttime,
         DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC,
           string Email_List_BCC)
        {
            //sttime = Convert.ToDateTime("2019-09-01");
            //ndtime= Convert.ToDateTime("2019-10-12");
            string dst = string.Empty;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("EnergyMeterReport_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Energy Meter Report generation has started...............");
                System.Data.DataTable dtEnergyData = AccessReportData.GetEnergyReportDataSona(sttime, ndtime, "S_GetSONA_EnergyMeterReport");
                bool Exported = false;
                if (dtEnergyData != null && dtEnergyData.Rows.Count > 0)
                {
                    Exported = Reports.EnergyMeterReport(dst, sttime, ndtime, dtEnergyData);
                }

                if (Exported)
                {
                    Logger.WriteDebugLog("Energy Meter Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("Energy Meter Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Sona Energy Report.");
                Logger.WriteErrorLog(ex.Message);
            }
        }

        internal static void ExportPlanVsActualReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, DateTime ndtime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string PlantID, string LineID)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("PlanVsActualReport_Tafe_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Plan Vs Actual Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                int Year = sttime.Year;
                int Month = sttime.Month;
                //string Date = Convert.ToDateTime(Year + "-" + Month + "-01").ToString("yyyy-MM-dd");
                string Date = Year + "-" + Month + "-01";
                System.Data.DataTable dtPlanVsActualDataCumulative = null;
                System.Data.DataTable dtPlanVsActualDataDaywise = AccessReportData.GetPlanVsActualData(PlantID, LineID, Date, out dtPlanVsActualDataCumulative);
                if (dtPlanVsActualDataDaywise != null && dtPlanVsActualDataDaywise.Rows.Count > 0)
                {
                    DataAvailable = true;
                    try
                    {
                        ExcelWorksheet worksheetCumulative = excelPackage.Workbook.Worksheets[1];
                        int rowStart = 8;
                        if (PlantID.Equals("All", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(PlantID))
                            worksheetCumulative.Cells["B4"].Value = "All";
                        else
                            worksheetCumulative.Cells["B4"].Value = PlantID;

                        if (LineID.Equals("All", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(LineID))
                            worksheetCumulative.Cells["E4"].Value = "All";
                        else
                            worksheetCumulative.Cells["E4"].Value = LineID;



                        //worksheetCumulative.Cells["C6"].Value = "MTD(" + sttime.ToString("MMMM") + ")";
                        //worksheetCumulative.Cells["I6"].Value = "YTD(" + sttime.ToString("yyyy") + ")";
                        worksheetCumulative.Cells["C6"].Value = "MTD(" + Convert.ToDateTime(dtPlanVsActualDataCumulative.Rows[0]["LastAggDateForMonth"].ToString()).ToString("dd-MMM-yyyy") + ")";
                        worksheetCumulative.Cells["I6"].Value = "YTD(" + Convert.ToDateTime(dtPlanVsActualDataCumulative.Rows[0]["LastAggDateForYear"].ToString()).ToString("dd-MMM-yyyy") + ")";

                        worksheetCumulative.Cells["G4"].Value = "As on date :";
                        worksheetCumulative.Cells["H4"].Value = sttime.ToString("yyyy-MM-dd");
                        foreach (DataRow dataRow in dtPlanVsActualDataCumulative.Rows)
                        {
                            worksheetCumulative.Cells[rowStart, 1].Value = dataRow["PartName"];
                            worksheetCumulative.Cells[rowStart, 2].Value = dataRow["PartID"];
                            worksheetCumulative.Cells[rowStart, 3].Value = dataRow["ScheduledQtyMTD"];
                            worksheetCumulative.Cells[rowStart, 4].Value = dataRow["ActualQtyMTD"];
                            worksheetCumulative.Cells[rowStart, 5].Value = dataRow["HoldQtyMTD"];
                            worksheetCumulative.Cells[rowStart, 6].Value = dataRow["DelayQtyMTD"];
                            worksheetCumulative.Cells[rowStart, 7].Value = dataRow["RejMaterialMTD"];
                            worksheetCumulative.Cells[rowStart, 8].Value = dataRow["RejProcessMTD"];
                            worksheetCumulative.Cells[rowStart, 9].Value = dataRow["ScheduledQtyYTD"];
                            worksheetCumulative.Cells[rowStart, 10].Value = dataRow["ActualQtyYTD"];
                            worksheetCumulative.Cells[rowStart, 11].Value = dataRow["HoldQtyYTD"];
                            worksheetCumulative.Cells[rowStart, 12].Value = dataRow["DelayQtyYTD"];
                            worksheetCumulative.Cells[rowStart, 13].Value = dataRow["RejMaterialYTD"];
                            worksheetCumulative.Cells[rowStart, 14].Value = dataRow["RejProcessYTD"];
                            rowStart++;
                        }
                        var worksheetDaywise = excelPackage.Workbook.Worksheets[2];
                        if (PlantID.Equals("All", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(PlantID))
                            worksheetDaywise.Cells["B4"].Value = "All";
                        else
                            worksheetDaywise.Cells["B4"].Value = PlantID;

                        if (LineID.Equals("All", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(LineID))
                            worksheetDaywise.Cells["E4"].Value = "All";
                        else
                            worksheetDaywise.Cells["E4"].Value = LineID;

                        rowStart = 8;
                        foreach (DataRow dataRow in dtPlanVsActualDataDaywise.Rows)
                        {
                            worksheetDaywise.Cells[rowStart, 1].Value = dataRow["Pdate"];
                            worksheetDaywise.Cells[rowStart, 1].Style.Numberformat.Format = "dd-MMM-yyyy";
                            worksheetDaywise.Cells[rowStart, 2].Value = dataRow["Line"];
                            worksheetDaywise.Cells[rowStart, 3].Value = dataRow["PartName"];
                            worksheetDaywise.Cells[rowStart, 4].Value = dataRow["PartID"];
                            worksheetDaywise.Cells[rowStart, 5].Value = dataRow["ScheduledQty"];
                            worksheetDaywise.Cells[rowStart, 6].Value = dataRow["ActualQty"];
                            worksheetDaywise.Cells[rowStart, 7].Value = dataRow["HoldQty"];
                            worksheetDaywise.Cells[rowStart, 8].Value = dataRow["DelayQty"];
                            worksheetDaywise.Cells[rowStart, 9].Value = dataRow["RejMaterial"];
                            worksheetDaywise.Cells[rowStart, 10].Value = dataRow["RejProcess"];
                            rowStart++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Plan Vs Actual Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("FPlan Vs Actual Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }

        internal static void ExportCategoryWiseOEEAndLossTimeReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string MachineID)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("OEEAndLosstimeDetails_Tafe_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("CategoryWise OEE and Loss Time Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                var worksheet = excelPackage.Workbook.Worksheets[1];
                worksheet.Cells["V1"].Value = sttime.ToString("dd-MM-yyyy");
                DataSet oeeAndLosstimeDetailsDataSet = new DataSet();
                try
                {
                    oeeAndLosstimeDetailsDataSet = AccessReportData.GetOEEAndLosstimeDetails(sttime, MachineID);
                    if (oeeAndLosstimeDetailsDataSet != null && oeeAndLosstimeDetailsDataSet.Tables.Count > 0)
                    {
                        //DataAvailable = true;
                        int CatNoProdCount = 0, CatBreakdownCount = 0, CatSpeedlossCount = 0;
                        int CategoryColNum = 4, RowNum = 4;
                        ExcelPackage Excel = new ExcelPackage(newFile, true);

                        System.Data.DataTable dtCategoryDetails = oeeAndLosstimeDetailsDataSet.Tables[0];
                        System.Data.DataTable dtDaywiseOeeAndLosstimeDetails = oeeAndLosstimeDetailsDataSet.Tables[1];
                        System.Data.DataTable dtDaywiseTotalOeeAndLosstimeDetails = oeeAndLosstimeDetailsDataSet.Tables[2];
                        System.Data.DataTable dtTotalOeeAndLosstimeDetails = oeeAndLosstimeDetailsDataSet.Tables[3];

                        if (dtCategoryDetails != null && dtCategoryDetails.Rows.Count > 0 && dtDaywiseOeeAndLosstimeDetails != null && dtDaywiseOeeAndLosstimeDetails.Rows.Count > 0 && dtDaywiseTotalOeeAndLosstimeDetails != null && dtDaywiseTotalOeeAndLosstimeDetails.Rows.Count > 0 && dtTotalOeeAndLosstimeDetails != null && dtTotalOeeAndLosstimeDetails.Rows.Count > 0) 
                            DataAvailable = true;

                        if (dtCategoryDetails != null && dtCategoryDetails.Rows.Count > 0)
                        {
                            List<string> listCatNoProdDetails = dtCategoryDetails.AsEnumerable().Where(x => x.Field<string>("MainCatagory").Equals("No production")).Select(x => x.Field<string>("Catagory")).ToList();
                            List<string> listCatBreakdownDetails = dtCategoryDetails.AsEnumerable().Where(x => x.Field<string>("MainCatagory").Equals("Breakdown")).Select(x => x.Field<string>("Catagory")).ToList();
                            List<string> listCatSpeedlossDetails = dtCategoryDetails.AsEnumerable().Where(x => x.Field<string>("MainCatagory").Equals("Speed loss")).Select(x => x.Field<string>("Catagory")).ToList();
                            CatNoProdCount = listCatNoProdDetails.Count;
                            CatBreakdownCount = listCatBreakdownDetails.Count;
                            CatSpeedlossCount = listCatSpeedlossDetails.Count;
                            if (CatNoProdCount > 0)
                            {
                                foreach (string category in listCatNoProdDetails)
                                {
                                    worksheet.Cells[RowNum, CategoryColNum].Value = category;
                                    CategoryColNum++;
                                }
                            }
                            CategoryColNum = 9;
                            if (CatBreakdownCount > 0)
                            {
                                foreach (string category in listCatBreakdownDetails)
                                {
                                    worksheet.Cells[RowNum, CategoryColNum].Value = category;
                                    CategoryColNum++;
                                }
                            }
                            CategoryColNum = 14;
                            if (CatSpeedlossCount > 0)
                            {
                                foreach (string category in listCatSpeedlossDetails)
                                {
                                    worksheet.Cells[RowNum, CategoryColNum].Value = category;
                                    CategoryColNum++;
                                }
                            }
                        }

                        RowNum = 5;
                        int ColNum = 1, CategoryCount = 1;
                        if (dtDaywiseOeeAndLosstimeDetails != null && dtDaywiseOeeAndLosstimeDetails.Rows.Count > 0)
                        {
                            foreach (DataRow dataRow in dtDaywiseOeeAndLosstimeDetails.Rows)
                            {
                                ColNum = 1; CategoryCount = 1;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["MachineID"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["AvlTotalTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["AvlTime"]; ColNum++;
                                for (int i = 1; i <= CatNoProdCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 8;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["LoadingTime"]; ColNum++;
                                for (int i = 1; i <= CatBreakdownCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 13;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["OperatingTime"]; ColNum++;
                                for (int i = 1; i <= CatSpeedlossCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 18;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["NetOperatingTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Hold"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["RejMat"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["RejPro"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["ValuableOperatingTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["AEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["PEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["QEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["OEffy"];
                                RowNum++;
                            }
                        }

                        if (dtDaywiseTotalOeeAndLosstimeDetails != null && dtDaywiseTotalOeeAndLosstimeDetails.Rows.Count > 0)
                        {
                            foreach (DataRow dataRow in dtDaywiseTotalOeeAndLosstimeDetails.Rows)
                            {
                                ColNum = 1; CategoryCount = 1;
                                worksheet.Cells[RowNum, ColNum].Value = "Total/Avgt"; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_AvlTotalTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_AvlTime"]; ColNum++;
                                for (int i = 1; i <= CatNoProdCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 8;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["LoadingTime"]; ColNum++;
                                for (int i = 1; i <= CatBreakdownCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 13;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["OperatingTime"]; ColNum++;
                                for (int i = 1; i <= CatSpeedlossCount; i++)
                                {
                                    worksheet.Cells[RowNum, ColNum].Value = Convert.ToDouble(dataRow["C" + CategoryCount]);
                                    ColNum++; CategoryCount++;
                                }
                                ColNum = 18;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["NetOperatingTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_Hold"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_RejMat"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_RejPro"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["ValuableOperatingTime"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_AEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_PEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_QEffy"]; ColNum++;
                                worksheet.Cells[RowNum, ColNum].Value = dataRow["Tot_OEffy"];
                                RowNum++;
                            }
                        }

                        RowNum = 17;
                        if (dtTotalOeeAndLosstimeDetails != null && dtTotalOeeAndLosstimeDetails.Rows.Count > 0)
                        {
                            foreach (DataRow dataRow in dtTotalOeeAndLosstimeDetails.Rows)
                            {
                                worksheet.Cells[RowNum, 1].Value = string.Format("Available Time: {0:0.##}% of total time", dataRow["AvailableTime"]).ToUpper();
                                worksheet.Cells[RowNum, 23].Value = string.Format("Plant Closure: {0:0.##}% of total time", dataRow["PlantClosureTime"]).ToUpper();
                                worksheet.Cells[RowNum + 1, 1].Value = string.Format("Loading Time: {0:0.##}% of total time", dataRow["LoadingTime"]).ToUpper();
                                //worksheet.Cells[RowNum + 1, 14].Value = string.Format("Others (P, A, M, RM): {0:0.##}% of total time", dataRow["Others"]).ToUpper();
                                worksheet.Cells[RowNum + 1, 18].Value = string.Format("No Prdn Planned: {0:0.##}% of total time", dataRow["NoPrdnPlanned"]).ToUpper();
                                worksheet.Cells[RowNum + 2, 1].Value = string.Format("Operating Time: {0:0.##}% of loading time", dataRow["OperatingTime"]).ToUpper();
                                worksheet.Cells[RowNum + 2, 10].Value = string.Format("Downtime: {0:0.##}% of Avl. Time", dataRow["DownTime"]).ToUpper();
                                worksheet.Cells[RowNum + 3, 1].Value = string.Format("Net Opt. Time: {0:0.##}% of loading Time", dataRow["NetOperatingTime"]).ToUpper();
                                worksheet.Cells[RowNum + 3, 8].Value = string.Format("Speed Loss: {0:0.##}% of Avl. Time", dataRow["SpeedLoss"]).ToUpper();
                                worksheet.Cells[RowNum + 4, 1].Value = string.Format("Val. Time: {0:0.##}% of loading Time", dataRow["ValuableOperatingTime"]).ToUpper();
                                worksheet.Cells[RowNum + 4, 6].Value = string.Format("Quality Loss: {0:0.##}% of Avl. Time", dataRow["QualityLoss"]).ToUpper();
                                RowNum++;
                            }
                        }

                        CategoryColNum = 4;
                        for (int i = CategoryColNum; i < CategoryColNum + 4; i++)
                        {
                            if (i >= CategoryColNum + CatNoProdCount)
                                worksheet.Column(i).Hidden = true;
                        }
                        CategoryColNum = 9;
                        for (int i = CategoryColNum; i < CategoryColNum + 4; i++)
                        {
                            if (i >= CategoryColNum + CatBreakdownCount)
                                worksheet.Column(i).Hidden = true;
                        }
                        CategoryColNum = 14;
                        for (int i = CategoryColNum; i < CategoryColNum + 4; i++)
                        {
                            if (i >= CategoryColNum + CatSpeedlossCount)
                                worksheet.Column(i).Hidden = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("CategoryWise OEE and Loss Time Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("CategoryWise OEE and Loss Time Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        internal static void ExportHoldReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, DateTime endtime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string LineID, string MachineID)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("HoldReport_Tafe_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Hold Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                var worksheet = excelPackage.Workbook.Worksheets[1];
                //worksheet.Cells["B4"].Value = sttime.ToString("yyyy-MM-dd");
                //worksheet.Cells["D4	"].Value = endtime.AddDays(-1).ToString("yyyy-MM-dd");
                System.Data.DataTable HoldReportDataTable = new System.Data.DataTable();
                try
                {
                    HoldReportDataTable = AccessReportData.GetHoldReportData(sttime, endtime, LineID, MachineID);
                    if (HoldReportDataTable != null && HoldReportDataTable.Rows.Count > 0)
                    {
                        DataAvailable = true;
                        int row = 7, col = 1;
                        worksheet.Cells["B4"].Value = sttime.ToString("yyyy-MM-dd");
                        worksheet.Cells["D4	"].Value = endtime.AddDays(-1).ToString("yyyy-MM-dd");
                        foreach (DataRow item in HoldReportDataTable.Rows)
                        {
                            col = 1;
                            worksheet.Cells[row, col].Value = item["PlantID"];
                            col++;
                            worksheet.Cells[row, col].Value = item["machineid"];
                            col++;
                            worksheet.Cells[row, col].Value = item["ShiftName"];
                            col++;
                            worksheet.Cells[row, col].Value = item["Employeeid"];
                            col++;
                            worksheet.Cells[row, col].Value = item["componentid"];
                            col++;
                            worksheet.Cells[row, col].Value = item["SupplierCode"];
                            col++;
                            worksheet.Cells[row, col].Value = item["HeatCode"];
                            col++;
                            worksheet.Cells[row, col].Value = item["BatchCode"];
                            col++;
                            worksheet.Cells[row, col].Value = item["description"];
                            col++;
                            worksheet.Cells[row, col].Value = item["compslno"];
                            col++;
                            worksheet.Cells[row, col].Value = item["QualityTS"];
                            col++;
                            worksheet.Cells[row, col].Value = item["Remark"];
                            col++;
                            worksheet.Cells[row, col].Value = item["OperatorRemarks"];
                            row++;
                        }
                        row--;
                        worksheet.Cells[4, 1, row, col + 1].AutoFitColumns();
                        worksheet.Cells[7, 1, row, col].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Hold Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("Hold Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        internal static void ExportMachineHistoryReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, DateTime endtime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string MachineID)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("MachineHistoryReport_Tafe_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("MachineHistory Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                var worksheet = excelPackage.Workbook.Worksheets[1];
                //worksheet.Cells["B4"].Value = sttime.ToString("yyyy-MM-dd");
                //worksheet.Cells["D4	"].Value = endtime.ToString("yyyy-MM-dd");
                List<MachineHistory> MachineHistoryData = new List<MachineHistory>();
                try
                {
                    MachineHistoryData = AccessReportData.GetMachineHistoryDatas(sttime, endtime, MachineID);
                    if (MachineHistoryData != null && MachineHistoryData.Count > 0)
                    {
                        DataAvailable = true;
                        int row = 7, col = 1;
                        worksheet.Cells["B4"].Value = sttime.ToString("yyyy-MM-dd");
                        worksheet.Cells["D4	"].Value = endtime.ToString("yyyy-MM-dd");
                        foreach (MachineHistory item in MachineHistoryData)
                        {
                            col = 1;
                            worksheet.Cells[row, col].Value = item.MachineID;
                            col++;
                            worksheet.Cells[row, col].Value = item.DownCode;
                            col++;
                            worksheet.Cells[row, col].Value = item.KindOfProblem;
                            col++;
                            worksheet.Cells[row, col].Value = item.Reason;
                            col++;
                            worksheet.Cells[row, col].Value = item.DownCategory;
                            col++;
                            worksheet.Cells[row, col].Value = item.BreakDownStart;
                            col++;
                            worksheet.Cells[row, col].Value = item.BreakDownEnd;
                            col++;
                            worksheet.Cells[row, col].Value = item.ActionToResolve;
                            col++;
                            worksheet.Cells[row, col].Value = item.ActionProposed;
                            col++;
                            worksheet.Cells[row, col].Value = item.TimeLost;
                            col++;
                            worksheet.Cells[row, col].Value = item.ElapsedTime;
                            col++;
                            worksheet.Cells[row, col].Value = item.Severity;
                            row++;
                        }
                        row--;
                        worksheet.Cells[4, 1, row, col + 1].AutoFitColumns();
                        worksheet.Cells[7, 1, row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[7, 1, row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[7, 1, row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[7, 1, row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[7, 1, row, col].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[7, 1, row, col].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[7, 1, row, col].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[7, 1, row, col].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);

                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("MachineHistory Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("MachineHistory Report not mailed: no data");
                }

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        internal static void ExportRejectionReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, DateTime endtime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string PlantID, string LineID, string Category)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }
                if (Category.Equals("Material", StringComparison.OrdinalIgnoreCase))
                    dst = Path.Combine(ExportPath, string.Format("TAFE_MaterialRejectionReport_{0:ddMMMyyyyHHmmss}.xlsx", sttime));
                else
                    dst = Path.Combine(ExportPath, string.Format("TAFE_ProcessRejectionReport_{0:ddMMMyyyyHHmmss}.xlsx", sttime));
                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Material Rejection Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                var worksheet = excelPackage.Workbook.Worksheets[1];
                worksheet.Cells["B4"].Value = sttime;

                worksheet.Cells["H4"].Value = LineID;
                worksheet.Cells["F4"].Value = PlantID;
                worksheet.Cells["J4"].Value = Category;
                PlantID = PlantID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : PlantID;
                LineID = LineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : LineID;
                System.Data.DataTable dtRejection = AccessReportData.GetRejectionReportData(sttime, endtime, PlantID, LineID, Category);

                try
                {
                    worksheet.Cells["D4"].Value = endtime.AddDays(-1);
                    int row = 9;
                    if (dtRejection != null && dtRejection.Rows.Count > 0)
                    {
                        DataAvailable = true;
                        foreach (DataRow dtRow in dtRejection.Rows)
                        {
                            worksheet.Cells[row, 1].Value = dtRow["Rejdate"];
                            worksheet.Cells[row, 2].Value = dtRow["ShiftName"];
                            worksheet.Cells[row, 3].Value = dtRow["Machineid"];
                            worksheet.Cells[row, 4].Value = dtRow["Employeeid"];
                            worksheet.Cells[row, 5].Value = dtRow["BatchCode"];
                            worksheet.Cells[row, 6].Value = dtRow["compslno"];
                            worksheet.Cells[row, 7].Value = dtRow["HeatCode"];
                            worksheet.Cells[row, 8].Value = dtRow["SupplierCode"];
                            worksheet.Cells[row, 9].Value = dtRow["componentid"];
                            worksheet.Cells[row, 10].Value = dtRow["description"];
                            worksheet.Cells[row, 11].Value = dtRow["Rejection_Qty"];
                            worksheet.Cells[row, 12].Value = dtRow["DefectObserved"];
                            worksheet.Cells[row, 13].Value = dtRow["MST"];
                            worksheet.Cells[row, 14].Value = dtRow["Remark"];
                            worksheet.Cells[row, 15].Value = dtRow["Scrap"];
                            worksheet.Cells[row, 16].Value = dtRow["Rew"];
                            worksheet.Cells[row, 17].Value = dtRow["Seg"];
                            worksheet.Cells[row, 18].Value = dtRow["AccUO"];
                            worksheet.Cells[row, 19].Value = dtRow["Rating"];
                            worksheet.Cells[row, 20].Value = dtRow["RootCause"];
                            worksheet.Cells[row, 21].Value = dtRow["ActiontoBeTaken"];
                            worksheet.Cells[row, 22].Value = dtRow["Targetdate"].ToString() != "" ? (DateTime.Parse(dtRow["Targetdate"].ToString()).Year == 1900 ? "" : dtRow["Targetdate"]) : "";
                            row++;
                        }
                        row--;
                        worksheet.Cells[4, 1, row, 24].AutoFitColumns();
                        worksheet.Cells[9, 1, row, 24].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[9, 1, row, 24].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[9, 1, row, 24].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[9, 1, row, 24].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[9, 1, row, 24].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[9, 1, row, 24].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[9, 1, row, 24].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[9, 1, row, 24].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Material Rejection Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("Material Rejection Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
        //internal static void ExportBatchwiseReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string PlantID, string LineID,string PartID, string Category)
        //{
        //	string dst = string.Empty;
        //	bool DataAvailable = false;
        //	try
        //	{
        //		if (!File.Exists(strReportFile))
        //		{
        //			Logger.WriteDebugLog("Template is not found on " + strReportFile);
        //			return;
        //		}

        //		if (!Directory.Exists(ExportPath))
        //		{
        //			Directory.CreateDirectory(ExportPath);
        //		}

        //		dst = Path.Combine(ExportPath, string.Format("Tafe_BatchWiseReport_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now));

        //		if (File.Exists(dst))
        //		{
        //			var di = new DirectoryInfo(dst);
        //			di.Attributes &= ~FileAttributes.ReadOnly;
        //			File.Delete(dst);
        //		}
        //		File.Copy(strReportFile, dst, true);
        //		Logger.WriteDebugLog("BatchWise Report generation has started...............");
        //		File.Copy(strReportFile, dst, true);
        //		FileInfo newFile = new FileInfo(dst);
        //		ExcelPackage pck = new ExcelPackage(newFile, true);
        //		var wsDts = pck.Workbook.Worksheets[1];
        //		wsDts.Cells["C5"].Value = sttime;
        //		wsDts.Cells["F5"].Value = PlantID;
        //		wsDts.Cells["J5"].Value = LineID;
        //		wsDts.Cells["N5"].Value = PartID;
        //		wsDts.Cells["R5"].Value = Category;
        //		PlantID = PlantID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : PlantID;
        //		LineID = LineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : LineID;
        //		System.Data.DataTable dtGraph = AccessReportData.GetBatchWiseGraphDateReport(sttime, PlantID, LineID, PartID, Category);
        //		System.Data.DataTable dtBatchwise = AccessReportData.GetBatchWiseDataReport(sttime, PlantID, LineID, PartID, Category);
        //		int graphrow = 10, datarow = 33;
        //		if (Category.Equals("Material", StringComparison.OrdinalIgnoreCase))
        //		{
        //			wsDts.Cells["A1"].Value = "BATCH WISE MATERIAL REJECTION REPORT";
        //		}
        //		else if (Category.Equals("Process", StringComparison.OrdinalIgnoreCase))
        //		{
        //			wsDts.Cells["A1"].Value = "BATCH WISE Process REJECTION REPORT";
        //		}
        //		string Partdescription = AccessReportData.Getdescription(PartID);
        //		if (dtGraph != null && dtGraph.Rows.Count > 0)
        //		{
        //			foreach (DataRow dtrow in dtGraph.Rows)
        //			{
        //				wsDts.Cells[graphrow, 30].Value = dtrow["Batchcode"];
        //				wsDts.Cells[graphrow, 31].Value = dtrow["RejectionPercent"];
        //				graphrow++;
        //			}


        //			var chart = (ExcelBarChart)wsDts.Drawings.AddChart("ColChart", eChartType.ColumnClustered);
        //			chart.SetSize(1180, 450);
        //			chart.SetPosition(5, 30, 0, 30);
        //			chart.Title.Text = "Batch Wise report";
        //			chart.XAxis.Title.Text = "BATCH";
        //			chart.YAxis.Title.Text = "REJECTION IN PERCENT";
        //			chart.Series.Add(ExcelRange.GetAddress(10, 31, graphrow - 1, 31), ExcelRange.GetAddress(10, 30, graphrow - 1, 30));
        //			chart.Series[0].Header = Partdescription;
        //		}
        //		try
        //		{
        //			bool exists = dtBatchwise.AsEnumerable().Where(c => c.Field<string>("Type").Equals("OK", StringComparison.OrdinalIgnoreCase)).Count() > 0;
        //			System.Data.DataTable dtOK = new System.Data.DataTable(); System.Data.DataTable dtRej = new System.Data.DataTable();
        //			if (exists)
        //				dtOK = dtBatchwise.Rows.Cast<DataRow>().Where(x => x["Type"].ToString().Equals("OK", StringComparison.OrdinalIgnoreCase)).CopyToDataTable();
        //			exists = dtBatchwise.AsEnumerable().Where(c => c.Field<string>("Type").Equals("Rejection", StringComparison.OrdinalIgnoreCase)).Count() > 0;
        //			if (exists)
        //				dtRej = dtBatchwise.Rows.Cast<DataRow>().Where(x => x["Type"].ToString().Equals("Rejection", StringComparison.OrdinalIgnoreCase)).CopyToDataTable();
        //			wsDts.Cells[datarow - 1, 2, datarow - 1, 4].Merge = true;
        //			wsDts.Cells[datarow - 1, 2].Value = "Part : " + Partdescription;
        //			wsDts.Cells[datarow - 1, 5].Value = "Total";
        //			int col = 5;
        //			for (int i = 5; i < dtBatchwise.Columns.Count; i++)
        //			{
        //				wsDts.Cells[32, (i + 1)].Value = dtBatchwise.Columns[i].ColumnName.ToString();
        //				col++;
        //			}
        //			wsDts.Cells[datarow - 1, 2, datarow - 1, (col)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //			wsDts.Cells[datarow - 1, 2, datarow - 1, (col)].Style.Font.Color.SetColor(Color.Blue);
        //			wsDts.Cells[datarow - 1, 2, datarow - 1, (col)].Style.Font.Bold = true;
        //			wsDts.Cells[datarow, 2].Value = "Status";
        //			wsDts.Cells[datarow, 3].Value = "OK";
        //			wsDts.Cells[datarow, 4].Value = "SupplierCode";
        //			wsDts.Cells[datarow, 2, datarow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //			wsDts.Cells[datarow, 2, datarow, 4].Style.Font.Bold = true;

        //			datarow++;
        //			if (dtOK.Rows.Count > 0)
        //			{
        //				foreach (DataRow dtrow in dtOK.Rows)
        //				{
        //					wsDts.Cells[datarow, 2].Value = dtrow["BatchStatus"];
        //					wsDts.Cells[datarow, 3].Value = dtrow["BatchCode"];
        //					wsDts.Cells[datarow, 4].Value = dtrow["Suppliercode"];
        //					wsDts.Cells[datarow, 5].Value = dtrow["TotalQty"];
        //					datarow++;
        //				}
        //				wsDts.Cells[datarow, 2, datarow, 4].Merge = true;
        //				wsDts.Cells[datarow, 2].Value = "Sum for status";
        //				wsDts.Cells[datarow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //				wsDts.Cells[datarow, 2].Style.Font.Bold = true;
        //				wsDts.Cells[datarow, 5].Formula = "=SUM(E34:E" + (datarow - 1) + ")";
        //				for (int k = 5; k < dtOK.Columns.Count; k++)
        //				{
        //					string Index = GetExcelColumnName(k + 1);
        //					string formula = "=SUM(" + Index + 34 + ":" + Index + (datarow - 1) + ")";
        //					wsDts.Cells[datarow, (k + 1)].Formula = formula;
        //				}
        //			}
        //			else
        //				wsDts.Cells[datarow, 5].Value = 0;
        //			int sumokrow = datarow;
        //			datarow = datarow + 2;
        //			wsDts.Cells[datarow, 2].Value = "Status";
        //			wsDts.Cells[datarow, 3].Value = "REJ";
        //			wsDts.Cells[datarow, 4].Value = "SupplierCode";
        //			wsDts.Cells[datarow, 2, datarow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //			wsDts.Cells[datarow, 2, datarow, 4].Style.Font.Bold = true;
        //			datarow++;
        //			int start = datarow;
        //			if (dtRej.Rows.Count > 0)
        //			{
        //				for (int i = 0; i < dtRej.Rows.Count; i++)
        //				{
        //					wsDts.Cells[datarow, 2].Value = dtRej.Rows[i]["BatchStatus"];
        //					wsDts.Cells[datarow, 3].Value = dtRej.Rows[i]["BatchCode"];
        //					wsDts.Cells[datarow, 4].Value = dtRej.Rows[i]["Suppliercode"];
        //					wsDts.Cells[datarow, 5].Value = dtRej.Rows[i]["TotalQty"];
        //					for (int k = 5; k < col; k++)
        //					{
        //						wsDts.Cells[datarow, (k + 1)].Value = string.IsNullOrEmpty(dtRej.Rows[i][k].ToString()) ? 0 : dtRej.Rows[i][k];
        //					}
        //					datarow++;
        //				}
        //				wsDts.Cells[datarow, 5].Formula = "=SUM(E" + start + ":E" + (datarow - 1) + ")";
        //			}
        //			else
        //				wsDts.Cells[datarow, 5].Value = 0;
        //			wsDts.Cells[datarow, 2, datarow, 4].Merge = true;
        //			wsDts.Cells[datarow, 2].Value = "Sum for status";


        //			for (int k = 5; k < dtRej.Columns.Count; k++)
        //			{
        //				string Index = GetExcelColumnName(k + 1);
        //				string formula = "=SUM(" + Index + start + ":" + Index + (datarow - 1) + ")";
        //				wsDts.Cells[datarow, (k + 1)].Formula = formula;
        //			}

        //			wsDts.Cells[datarow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //			wsDts.Cells[datarow, 2].Style.Font.Bold = true;
        //			int sumrejrow = datarow;
        //			datarow = datarow + 2;
        //			wsDts.Cells[datarow, 2, datarow, 4].Merge = true;
        //			wsDts.Cells[datarow, 2].Value = "Sum for Parts for Month";
        //			wsDts.Cells[datarow, 5].Formula = "=SUM(E" + sumokrow + ",E" + sumrejrow + ")";
        //			for (int k = 5; k < dtRej.Columns.Count; k++)
        //			{
        //				string Index = GetExcelColumnName(k + 1);
        //				string formula = "=SUM(" + Index + sumokrow + "," + Index + sumrejrow + ")";
        //				wsDts.Cells[datarow, (k + 1)].Formula = formula;
        //			}
        //			wsDts.Cells[datarow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //			wsDts.Cells[datarow, 2].Style.Font.Bold = true;
        //			wsDts.Cells[32, 1, datarow, col].AutoFitColumns();
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
        //			wsDts.Cells[33, 2, datarow, col].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
        //		}
        //		catch (Exception ex)
        //		{
        //			Logger.WriteErrorLog(ex.ToString());
        //		}
        //		if (DataAvailable)
        //		{
        //			pck.SaveAs(newFile);
        //			Logger.WriteDebugLog("BatchWise Report Exported successfully.");
        //			SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
        //		}
        //		else
        //		{
        //			Logger.WriteDebugLog("BatchWise Report not mailed: no data");
        //		}
        //	}
        //	catch (Exception ex)
        //	{
        //		Logger.WriteErrorLog(ex.ToString());
        //	}
        //}
        internal static void ExportLineMeterReport(string strReportFile, string ExportPath, string ExportedReportFile, int ExportType, DateTime sttime, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string PlantID, string LineID)
        {
            string dst = string.Empty;
            bool DataAvailable = false;
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("LineMeterReportTafe_{0:ddMMMyyyyHHmmss}.xlsx", sttime));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                string Fromdate = AccessReportData.GetLogicalMonthStartEnd(sttime, "start"); ;
                string toDate = AccessReportData.GetLogicalMonthStartEnd(sttime, "end");
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Line Meter Report generation has started...............");
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                ws.Cells["A1"].Value = "Date: " + sttime.ToString("dd-MM-yyyy");
                ws.Cells["Z1"].Value = "Line-Id";
                ws.Cells["AC1"].Value = LineID;
                System.Data.DataTable dt = AccessReportData.GetLinemeterData(LineID, Convert.ToDateTime(Fromdate).ToString("yyyy-MM-dd HH:mm:ss"), Convert.ToDateTime(toDate).ToString("yyyy-MM-dd HH:mm:ss"));
                if (dt != null && dt.Rows.Count > 0)
                {
                    DataAvailable = true;

                    int col = 2;
                    foreach (DataRow Row in dt.Rows)
                    {
                        ws.Cells[3, col].Value = Convert.ToDateTime(Row["Day"].ToString());
                        ws.Cells[4, col].Value = Convert.ToDouble(Row["TargetCount"].ToString());
                        ws.Cells[5, col].Value = Convert.ToDouble(Row["ActualCount"].ToString());
                        ws.Cells[6, col].Value = Convert.ToDouble(Row["DelayCount"].ToString());
                        ws.Cells[7, col].Value = Convert.ToDouble(Row["TenPercent"].ToString());
                        ws.Cells[8, col].Value = Convert.ToDouble(Row["NegativeTenPercent"].ToString());
                        ws.Cells[9, col].Value = Convert.ToDouble(Row["LoadingHours"].ToString());
                        ws.Cells[10, col].Value = Convert.ToDouble(Row["NoOfManpower"].ToString());
                        ws.Cells[12, col].Value = Convert.ToDouble(Row["Okdays"].ToString());
                        if ((Convert.ToDouble(Row["Okdays"].ToString())).Equals(1))
                        {
                            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B050");
                            ws.Cells[5, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[5, col].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        ws.Cells[11, col].Value = Convert.ToDouble(Row["LineEfficiency"].ToString());
                        col++;
                    }
                    
                    for (int days = dt.Rows.Count + 1; days <= 31; days++)
                    {
                        ws.Column(days + 1).Hidden = true;

                    }

                    ws.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    ws.Cells["AG4"].Calculate();
                    ws.Cells["AG5"].Calculate();
                    ws.Cells["AG6"].Calculate();
                    ws.Cells["AG7"].Calculate();
                    ws.Cells["AG8"].Calculate();
                    ws.Cells["AG12"].Calculate();
                    ws.Cells["AH12"].Calculate();
                    ws.Cells["AI12"].Calculate();
                    ws.Cells["AI5"].Calculate();
                    ws.Cells["AJ5"].Calculate();
                    ws.Cells["AK5"].Calculate();
                }
                if (DataAvailable)
                {
                    excelPackage.SaveAs(newFile);
                    Logger.WriteDebugLog("Line Meter Report Exported successfully.");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("Line Meter Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }

        public static void ExportDailyProdDownDaywiseExcelReportOnlyLnT(string strtTime, string endTime, string strReportFile, string ExportPath,
            string ExportedReportFile, string MachineId, string operators, string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            string dst = string.Empty;
            //strtTime = "2019-11-13"; // g: test 
            //endTime = "2019-11-14"; // g: test 
            Dictionary<string, int> dctrowsPrev = new Dictionary<string, int>();
            strtTime = AccessReportData.Gellogicalmonthstart(Convert.ToDateTime(strtTime));
            //endTime = AccessReportData.GetLogicalDayEnd(DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                Logger.WriteDebugLog(string.Format("Start Time={0}, End Time={1}", strtTime, endTime));
                bool dataAvailable = false;
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SM_CumulativeProductionandDowntimeDetails_{1}_{0:ddMMMyyyyHHmmss}.xlsx",Convert.ToDateTime(endTime), isDayWise ? "Day" : "Shift"));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                File.Copy(strReportFile, dst, true);
                Logger.WriteDebugLog("Template Copied to Export Path");
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                ws.Cells["B3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm tt");
                ws.Cells["D3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm tt");

                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    MachineId = "";
                }

                if (plantid.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    plantid = "";
                }
                System.Data.DataTable dtMachinelist;
                string parameter = "Summary";
                System.Data.DataTable dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter, out dtMachinelist);
                if (dt != null && dt.Rows.Count > 0) dataAvailable = true;
                Logger.WriteDebugLog("Values obtain for Parameter- Summary");
                int r = 7;
                Dictionary<string, int> dctrows = new Dictionary<string, int>();
                foreach (DataRow row in dtMachinelist.Rows)
                {
                    if (r == 7)
                    {
                        ws.Name = row["Machineid"].ToString();
                    }
                    else
                    {
                        excelPackage.Workbook.Worksheets.Add(row["Machineid"].ToString(), ws);
                    }
                    ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                    ws.Cells["F3"].Value = plantid.Equals("") ? "ALL" : plantid;
                    ws.Cells["H3"].Value = row["Machineid"].ToString();

                    dctrows.Add(row["Machineid"].ToString(), 9);
                    r++;
                }

                r = 7;
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                        //ws.Cells["F3"].Value = plantid.Equals("") ? "ALL" : plantid;
                        //ws.Cells["H3"].Value = row["Machineid"].ToString();
                        ws.Cells[r, 1].Value = row["Totaltime"];
                        ws.Cells[r, 2].Value = row["AvailableTime"];
                        ws.Cells[r, 3].Value = row["Runtime"];
                        ws.Cells[r, 4].Value = row["NetDowntime"];
                        ws.Cells[r, 5].Value = row["PDT"];
                        ws.Cells[r, 6].Value = row["RuntimeEffy"];
                        //ws.Cells[r, 5].Value = row["Managementloss"];
                        //ws.Cells[6, 1, r + 1, 6].AutoFitColumns();
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                        // ignore machine if not in summary
                    }
                }


                //parameter = "Efficiency";
                //dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, MachineId, plantid, parameter);
                //r = 10;
                //foreach (DataRow row in dt.Rows)
                //{
                //    try
                //    {
                //        ws = excelPackage.Workbook.Worksheets[row["Machineid"].ToString()];
                //        ws.Cells[r, 2].Value = row["CycleCount"];
                //        ws.Cells[r, 4].Value = row["ProductionEfficiency"];
                //        ws.Cells[r, 6].Value = row["AvailabilityEfficiency"];
                //        ws.Cells[r, 8].Value = row["OverAllEfficiency"];
                //    }
                //    catch (Exception ex)
                //    {
                //        // ignore machine if not in summary
                //    }
                //}

                #region "DownTime Summary"

                parameter = "DowntimeSummary";
                //'2019-10-01 06:00:00','2019-10-02 06:00:00'
                foreach (string curMachine in dctrows.Keys.ToList())
                {
                    try
                    {
                        ws = excelPackage.Workbook.Worksheets[curMachine];
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "Downtime Summary:";
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[dctrows[curMachine], 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        dctrows[curMachine] += 1;
                        ws.Cells[dctrows[curMachine], 1].Value = "DownDescription";
                        ws.Cells[dctrows[curMachine], 2].Value = "Sum Downtime in min";
                        ws.Cells[dctrows[curMachine], 3].Value = "No of Occurences";
                        ws.Cells[dctrows[curMachine], 4].Value = "MIN. Downtime in min";
                        ws.Cells[dctrows[curMachine], 5].Value = "MAX. Downtime in min";
                        ws.Cells[dctrows[curMachine], 6].Value = "DownTime %";
                        ws.Cells[dctrows[curMachine] - 1, 1, dctrows[curMachine], 6].Style.Font.Bold = true;
                        dctrows[curMachine] += 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }

                foreach (string machine in dctrows.Keys.ToList())
                {
                    r = 12;
                    ws = excelPackage.Workbook.Worksheets[machine];
                    dt = AccessReportData.GetProductionAndDowntimes(strtTime, endTime, machine, plantid, parameter, out dtMachinelist);
                    Logger.WriteDebugLog("Values obtain for Parameter- DowntimeSummary");
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            ws.Cells[dctrows[machine], 1].Value = row["DownDescription"];
                            ws.Cells[dctrows[machine], 2].Value = row["DownTime"];
                            ws.Cells[dctrows[machine], 3].Value = Convert.ToInt32(row["NoOfOccurences"]);
                            ws.Cells[dctrows[machine], 4].Value = row["MinDowntime"];
                            ws.Cells[dctrows[machine], 5].Value = row["MaxDowntime"];
                            ws.Cells[dctrows[machine], 6].Value = Convert.ToDouble(row["DowntimePercent"]);
                            dctrows[machine] += 1;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }
                }
                foreach (string machine in dctrows.Keys.ToList())
                {
                    ws = excelPackage.Workbook.Worksheets[machine];

                    using (ExcelRange range = ws.Cells[10, 1, dctrows[machine] - 1, 6])
                    {
                        range.AutoFitColumns();
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    }
                    dctrowsPrev[machine] = dctrows[machine];
                }
                #endregion

                excelPackage.SaveAs(newFile);

                if (dataAvailable)
                {
                    Logger.WriteDebugLog("CumulativeProductionandDowntime Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("CumulativeProductionandDowntime Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error: " + ex.ToString());
                Logger.WriteErrorLog(ex.StackTrace);
            }
        }

        #region Maintenance Checklist Report - GEA
        internal static void GenerateWeeklyChklistReport(string strReportFile, string ExportPath, string ExportedReportFile, string machineID, string LineId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            bool dataAvailable = false;
            try
            {
                string dst = string.Empty;
                dst = Path.Combine(ExportPath, string.Format("MaintenanceCheckList_GEA_Weekly_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Yearly Maintenance Checklist Report template does not exists at - " + strReportFile);
                    dataAvailable = false;
                }
                else
                {
                    int rowStart = 8;
                    int colStart = 5;
                    int data_row_num = 1;
                    File.Copy(strReportFile, dst, true);
                    FileInfo newFile = new FileInfo(dst);
                    ExcelPackage Excel = new ExcelPackage(newFile, true);
                    var wrkshtYearlyMaintenanceChklist = Excel.Workbook.Worksheets[1];
                    wrkshtYearlyMaintenanceChklist.Cells["B4"].Value = LineId;
                    wrkshtYearlyMaintenanceChklist.Cells["AS3"].Value = "Date : " + DateTime.Now.ToString("dd-MMM-yyyy");
                    System.Data.DataTable dtWeeklyChklistReportData = AccessReportData.GetWeeklyChklistReportData(machineID, DateTime.Now.Year);
                    
                    if (dtWeeklyChklistReportData != null && dtWeeklyChklistReportData.Rows.Count > 0 && dtWeeklyChklistReportData.Columns.Count > 7)
                    {
                        dataAvailable = true;
                        var ListOfMachines = dtWeeklyChklistReportData.AsEnumerable().Select(s => s.Field<string>("Machineid")).Distinct();
                        List<string> spanHeaders = dtWeeklyChklistReportData.Columns.Cast<DataColumn>().Where(x => x.Ordinal > 6).Select(x => x.ColumnName.Contains("-") ? x.ColumnName.Split('-')[0] : x.ColumnName).ToList();
                        List<string> weeksList = dtWeeklyChklistReportData.Columns.Cast<DataColumn>().Where(x => x.Ordinal > 6).Select(x => x.ColumnName.Contains("-") ? x.ColumnName.Split('-')[1] : x.ColumnName).ToList();
                        Dictionary<string, int> headerCounts = spanHeaders.GroupBy(x => x).Select(x => new { Month = x.Key, Count = x.Count() }).ToDictionary(x => x.Month, x => x.Count);

                        foreach (KeyValuePair<string, int> keyValuePair in headerCounts)
                        {
                            wrkshtYearlyMaintenanceChklist.Cells[6, colStart].Value = keyValuePair.Key;
                            wrkshtYearlyMaintenanceChklist.Cells[6, colStart, 6, colStart + (keyValuePair.Value - 1)].Merge = true;
                            wrkshtYearlyMaintenanceChklist.Cells[6, colStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colStart += keyValuePair.Value;
                        }

                        colStart = 5;
                        foreach (string weekName in weeksList)
                        {
                            wrkshtYearlyMaintenanceChklist.Cells[7, colStart].Value = weekName;
                            colStart++;
                        }

                        int r = 1;
                        foreach (string machine in ListOfMachines)
                        {
                            if (r == 1)
                            {
                                wrkshtYearlyMaintenanceChklist.Name = machine;
                            }
                            else
                            {
                                Excel.Workbook.Worksheets.Add(machine, wrkshtYearlyMaintenanceChklist);
                            }
                            r++;
                        }
                        foreach(string machine in ListOfMachines)
                        {
                            rowStart = 8;
                            data_row_num = 1;
                            wrkshtYearlyMaintenanceChklist = Excel.Workbook.Worksheets[machine];
                            wrkshtYearlyMaintenanceChklist.Cells["D4"].Value = machine;
                            System.Data.DataTable dtMacWeeklyChklistReportData = dtWeeklyChklistReportData.AsEnumerable().Where(x => x.Field<string>("Machineid").Equals(machine)).CopyToDataTable();

                            for (int i = 0; i < dtMacWeeklyChklistReportData.Rows.Count; i++)
                            {
                                colStart = 5;
                                wrkshtYearlyMaintenanceChklist.Cells[rowStart, 1].Value = data_row_num;
                                wrkshtYearlyMaintenanceChklist.Cells[rowStart, 2].Value = dtMacWeeklyChklistReportData.Rows[i]["Chekpoints"];
                                wrkshtYearlyMaintenanceChklist.Cells[rowStart, 3].Value = dtMacWeeklyChklistReportData.Rows[i]["Method"];
                                wrkshtYearlyMaintenanceChklist.Cells[rowStart, 4].Value = dtMacWeeklyChklistReportData.Rows[i]["Criteria"];
                                for (int j = 7; j < dtMacWeeklyChklistReportData.Columns.Count; j++)
                                {
                                    if (dtMacWeeklyChklistReportData.Rows[i][j].ToString().Equals("1"))
                                    {
                                        var circle_done = wrkshtYearlyMaintenanceChklist.Drawings.AddShape("Circle_ActDone" + DateTime.Now.ToString("ddMMyyyyhhmmfff"), eShapeStyle.FlowChartConnector);
                                        circle_done.Fill.Style = eFillStyle.SolidFill;
                                        circle_done.Fill.Transparancy = 20;
                                        circle_done.Border.Fill.Style = eFillStyle.SolidFill;
                                        circle_done.Border.LineStyle = eLineStyle.Solid;
                                        circle_done.Border.Width = 1;
                                        circle_done.Border.Fill.Color = Color.Black;
                                        circle_done.Border.LineCap = eLineCap.Round;
                                        circle_done.Fill.Color = Color.Green;
                                        circle_done.SetSize(14, 14);
                                        circle_done.SetPosition(rowStart - 1, 3, colStart, 3);
                                    }
                                    if (dtMacWeeklyChklistReportData.Rows[i][j].ToString().Equals("2"))
                                    {
                                        var circle_notdone = wrkshtYearlyMaintenanceChklist.Drawings.AddShape("Circle_ActNotDone" + DateTime.Now.ToString("ddMMyyyyhhmmfff"), eShapeStyle.FlowChartConnector);
                                        circle_notdone.Fill.Style = eFillStyle.SolidFill;
                                        circle_notdone.Fill.Transparancy = 20;
                                        circle_notdone.Border.Fill.Style = eFillStyle.SolidFill;
                                        circle_notdone.Border.LineStyle = eLineStyle.Solid;
                                        circle_notdone.Border.Width = 1;
                                        circle_notdone.Border.Fill.Color = Color.Black;
                                        circle_notdone.Border.LineCap = eLineCap.Round;
                                        circle_notdone.Fill.Color = Color.Red;
                                        circle_notdone.SetSize(14, 14);
                                        circle_notdone.SetPosition(rowStart - 1, 3, colStart, 3);
                                    }
                                    if (dtMacWeeklyChklistReportData.Rows[i][j].ToString().Equals("3"))
                                    {
                                        var circle_chkdone_replaced = wrkshtYearlyMaintenanceChklist.Drawings.AddShape("Circle_ActChkDoneReplaced" + DateTime.Now.ToString("ddMMyyyyhhmmfff"), eShapeStyle.FlowChartConnector);
                                        circle_chkdone_replaced.Fill.Style = eFillStyle.SolidFill;
                                        circle_chkdone_replaced.Fill.Transparancy = 20;
                                        circle_chkdone_replaced.Border.Fill.Style = eFillStyle.SolidFill;
                                        circle_chkdone_replaced.Border.LineStyle = eLineStyle.Solid;
                                        circle_chkdone_replaced.Border.Width = 1;
                                        circle_chkdone_replaced.Border.Fill.Color = Color.Black;
                                        circle_chkdone_replaced.Border.LineCap = eLineCap.Round;
                                        circle_chkdone_replaced.Fill.Color = Color.Blue;
                                        circle_chkdone_replaced.SetSize(14, 14);
                                        circle_chkdone_replaced.SetPosition(rowStart - 1, 3, colStart, 3);
                                    }
                                    colStart++;
                                }
                                rowStart++;
                                data_row_num++;
                            }
                            rowStart--;
                            //wrkshtYearlyMaintenanceChklist.Cells[6, 1, rowStart, 66].AutoFitColumns();
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtYearlyMaintenanceChklist.Cells[8, 1, rowStart, 66].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        }
                        
                    }
                    Excel.SaveAs(newFile);
                    //DownloadMultipleFile(dst, Excel.GetAsByteArray());
                    if (dataAvailable)
                    {
                        Logger.WriteDebugLog("GEA_YearlyMaintenanceChecklist Report Exported successfully");
                        SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                    }
                    else
                    {
                        Logger.WriteDebugLog("GEA_YearlyMaintenanceChecklist Report not mailed: no data");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
                throw;
            }
        }

        internal static void GenerateDailyChklistReport(string strReportFile, string ExportPath, string ExportedReportFile, string machineID, string LineID, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, bool isDayWise)
        {
            bool dataAvailable = false;
            string startdate = AccessReportData.GetLogicalMonthStartEnd(DateTime.Now, "Start");
            try
            {
                string dst = string.Empty;
                dst = Path.Combine(ExportPath, string.Format("MaintenanceCheckList_GEA_Daywise_{0:ddMMMyyyyHHmmss}.xlsx", DateTime.Now));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Daily Maintenance Checklist Report template does not exists at - " + strReportFile);
                    dataAvailable = false;
                }
                else
                {
                    int rowStart = 7;
                    int colStart = 5;
                    int data_row_num = 1;
                    File.Copy(strReportFile, dst, true);
                    FileInfo newFile = new FileInfo(dst);
                    ExcelPackage Excel = new ExcelPackage(newFile, true);
                    var wrkshtDailyMaintenanceChklist = Excel.Workbook.Worksheets[1];
                    wrkshtDailyMaintenanceChklist.Cells["B4"].Value = LineID;
                    
                    wrkshtDailyMaintenanceChklist.Cells["I4"].Value = "Month : " + DateTime.Now.ToString("MMM");
                    wrkshtDailyMaintenanceChklist.Cells["O4"].Value = "Year : " + DateTime.Now.Year;
                    System.Data.DataTable dtDailyChklistReportData = AccessReportData.GetDailyChecklistReportData(LineID, machineID, startdate);

                    if (dtDailyChklistReportData != null && dtDailyChklistReportData.Rows.Count > 0 && dtDailyChklistReportData.Columns.Count > 7)
                    {
                        dataAvailable = true;
                        var ListOfMachines = dtDailyChklistReportData.AsEnumerable().Select(s => s.Field<string>("Machineid")).Distinct();
                        List<string> spanHeaders = dtDailyChklistReportData.Columns.Cast<DataColumn>().Where(x => x.Ordinal > 6).Select(x => x.ColumnName.Contains("-") ? x.ColumnName.Split('-')[2] : x.ColumnName).ToList();

                        foreach(string header in spanHeaders)
                        {
                            wrkshtDailyMaintenanceChklist.Cells[6, colStart].Value = Convert.ToInt32(header);
                            colStart++;
                        }
                        wrkshtDailyMaintenanceChklist.Cells[6, colStart].Value = "Remarks";
                        wrkshtDailyMaintenanceChklist.Cells[6, colStart].AutoFitColumns();
                        int r = 1;
                        foreach (string machine in ListOfMachines)
                        {
                            if (r == 1)
                            {
                                wrkshtDailyMaintenanceChklist.Name = machine;
                            }
                            else
                            {
                                Excel.Workbook.Worksheets.Add(machine, wrkshtDailyMaintenanceChklist);
                            }
                            r++;
                        }
                        foreach (string machine in ListOfMachines)
                        {
                            rowStart = 7;
                            data_row_num = 1;
                            wrkshtDailyMaintenanceChklist = Excel.Workbook.Worksheets[machine];
                            wrkshtDailyMaintenanceChklist.Cells["C4"].Value = "Machine Name : " + machine;
                            System.Data.DataTable dtMacWeeklyChklistReportData = dtDailyChklistReportData.AsEnumerable().Where(x => x.Field<string>("Machineid").Equals(machine)).CopyToDataTable();

                            for (int i = 0; i < dtMacWeeklyChklistReportData.Rows.Count; i++)
                            {
                                colStart = 5;
                                wrkshtDailyMaintenanceChklist.Cells[rowStart, 1].Value = data_row_num;
                                wrkshtDailyMaintenanceChklist.Cells[rowStart, 2].Value = dtMacWeeklyChklistReportData.Rows[i]["Activity"];
                                wrkshtDailyMaintenanceChklist.Cells[rowStart, 3].Value = dtMacWeeklyChklistReportData.Rows[i]["Method"];
                                wrkshtDailyMaintenanceChklist.Cells[rowStart, 4].Value = dtMacWeeklyChklistReportData.Rows[i]["Criteria"];
                                for (int j = 7; j < dtMacWeeklyChklistReportData.Columns.Count; j++)
                                {
                                    if(!string.IsNullOrEmpty(dtMacWeeklyChklistReportData.Rows[i][j].ToString()))
                                        wrkshtDailyMaintenanceChklist.Cells[rowStart, colStart].Value = dtMacWeeklyChklistReportData.Rows[i][j].ToString();
                                    colStart++;
                                }
                                rowStart++;
                                data_row_num++;
                            }
                            rowStart--;
                            wrkshtDailyMaintenanceChklist.Cells[6, 1, rowStart, 4].AutoFitColumns();
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            wrkshtDailyMaintenanceChklist.Cells[7, 1, rowStart, 36].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        }

                    }
                    Excel.SaveAs(newFile);
                    if (dataAvailable)
                    {
                        Logger.WriteDebugLog("DailyMaintenanceChecklist_GEA Report Exported successfully");
                        SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                    }
                    else
                    {
                        Logger.WriteDebugLog("GEA_YearlyMaintenanceChecklist Report not mailed: no data");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
                throw;
            }
        }
        #endregion

        #region Production Details Report - LnT
        internal static void ExportLnTProductionDetailsReport(DateTime strtTime, DateTime endTime, string strReportFile, string ExportPath, string ExportedReportFile, string MachineId, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string RunReportForEvery)
        {
            string dst = string.Empty;
            bool dataAvailable = false;
            
            try
            {
                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("{1}_{2}_{0:ddMMMyyyyHHmmss}.xlsx", endTime,ExportedReportFile,RunReportForEvery));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }
                System.Data.DataTable dt = AccessReportData.GetComponentDetailsReport(strtTime, endTime, MachineId);
                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage excelpackage = new ExcelPackage(newFile, true);
                ExcelWorksheet ws = excelpackage.Workbook.Worksheets[1];
                ws.Cells["B4"].Value = strtTime.ToString("dd-MMM-yyyy HH:mm:ss");
                ws.Cells["D4"].Value = endTime.ToString("dd-MMM-yyyy HH:mm:ss");
                
                if (dt != null && dt.Rows.Count > 0)
                {
                    dataAvailable = true;
                    var mac = dt.AsEnumerable().Select(s => s.Field<string>("Machineid")).Distinct();
                    int row = 10;
                    int r = 1;
                    foreach(string machine in mac)
                    {
                        if (r == 1)
                        {
                            ws.Name = machine;
                        }
                        else
                        {
                            excelpackage.Workbook.Worksheets.Add(machine, ws);
                        }
                        r++;
                    }
                    foreach(string machine in mac)
                    {
                        try
                        {
                            int rowNumber = 1;
                            row = 10;
                            ws = excelpackage.Workbook.Worksheets[machine];
                            ws.Cells["B5"].Value = machine;
                            System.Data.DataTable dtMac = dt.AsEnumerable().Where(x => x.Field<string>("Machineid").Equals(machine)).CopyToDataTable();
                            //dt = AccessReportData.GetComponentDetailsReport(strtTime, endTime, machine);
                            foreach(DataRow Dtrow in dtMac.Rows)
                            {
                                ws.Cells[row, 1].Value = rowNumber;
                                ws.Cells[row, 2].Value = DBNull.Value.Equals(Dtrow["Component"])? string.Empty: Dtrow["Component"].ToString();
                                ws.Cells[row, 3].Value = DBNull.Value.Equals(Dtrow["CompDescription"]) ? string.Empty : Dtrow["CompDescription"].ToString();
                                ws.Cells[row, 4].Value = DBNull.Value.Equals(Dtrow["Operation"]) ? string.Empty : Dtrow["Operation"].ToString();
                                ws.Cells[row, 5].Value = DBNull.Value.Equals(Dtrow["ActualQty"]) ? 0 : Convert.ToDouble(Dtrow["ActualQty"].ToString());
                                ws.Cells[row, 6].Value = DBNull.Value.Equals(Dtrow["ActualCycletime"]) ? string.Empty : Dtrow["ActualCycletime"].ToString();
                                ws.Cells[row, 7].Value = DBNull.Value.Equals(Dtrow["Remarks"]) ? string.Empty : Dtrow["Remarks"].ToString();
                                ws.Cells[row, 8].Value = DBNull.Value.Equals(Dtrow["BatchStart"]) ? string.Empty : Dtrow["BatchStart"].ToString();
                                ws.Cells[row, 9].Value = DBNull.Value.Equals(Dtrow["BatchEnd"]) ? string.Empty : Dtrow["BatchEnd"].ToString();
                                //ws.Cells[row, 7].Value = Dtrow["ActualQty"];
                                ws.Cells[row, 10].Value = DBNull.Value.Equals(Dtrow["TargetCycleTime"]) ? string.Empty : Dtrow["TargetCycleTime"].ToString();
                                ws.Cells[row, 11].Value = DBNull.Value.Equals(Dtrow["Diff"]) ? string.Empty : Dtrow["Diff"].ToString();
                                //ws.Cells[row, 9].Value = Dtrow["ActualCycletime"];
                                //ws.Cells[row, 10].Value = Dtrow["PDT"];
                                row++;
                                rowNumber++;
                            }
                            row--;
                            ws.Cells[2, 1, row, 11].AutoFitColumns();
                            ws.Cells[10, 1, row, 11].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[10, 1, row, 11].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[10, 1, row, 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[10, 1, row, 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[10, 1, row, 11].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            ws.Cells[10, 1, row, 11].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            ws.Cells[10, 1, row, 11].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            ws.Cells[10, 1, row, 11].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            ws.Cells[10, 1, row, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[10, 1, row, 11].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));
                            ws.Cells[10, 8, row, 9].Style.Fill.BackgroundColor.SetColor(Color.White);
                            ws.Cells[10, 1, row, 11].Style.Font.Color.SetColor(Color.Black);
                            ws.Column(6).Width = 19;
                            ws.Column(10).Width = 19;
                            ws.Column(11).Width = 19;
                        }
                        catch(Exception ex)
                        {
                            Logger.WriteErrorLog(ex.Message);
                        }
                        
                    }
                    //foreach (DataRow Dtrow in dt.Rows)
                    //{
                    //    ws = excelpackage.Workbook.Worksheets[Dtrow["Machineid"].ToString()];
                    //    //ws.Cells[row, 1].Value = Dtrow["Machineid"];
                    //    ws.Cells[row, 1].Value = Dtrow["Component"];
                    //    //ws.Cells[row, 3].Value = Dtrow["CompDescription"];
                    //    ws.Cells[row, 2].Value = Dtrow["Operation"];
                    //    ws.Cells[row, 3].Value = Dtrow["ActualQty"];
                    //    ws.Cells[row, 4].Value = Dtrow["ActualCycletime"];
                    //    ws.Cells[row, 5].Value = Dtrow["Remarks"];
                    //    ws.Cells[row, 6].Value = Dtrow["BatchStart"];
                    //    ws.Cells[row, 7].Value = Dtrow["BatchEnd"];
                    //    //ws.Cells[row, 7].Value = Dtrow["ActualQty"];
                    //    ws.Cells[row, 8].Value = Dtrow["TargetCycleTime"];
                    //    ws.Cells[row, 9].Value = Dtrow["Diff"];
                    //    //ws.Cells[row, 9].Value = Dtrow["ActualCycletime"];
                    //    //ws.Cells[row, 10].Value = Dtrow["PDT"];
                    //    row++;
                    //}
                    //row--;
                    //ws.Cells[2, 1, row, 9].AutoFitColumns();
                    //ws.Cells[10, 1, row, 9].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[10, 1, row, 9].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[10, 1, row, 9].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[10, 1, row, 9].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[10, 1, row, 9].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Cells[10, 1, row, 9].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Cells[10, 1, row, 9].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Cells[10, 1, row, 9].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    //ws.Cells[10, 1, row, 9].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells[10, 1, row, 9].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));
                }
                excelpackage.SaveAs(newFile);
                if (dataAvailable)
                {
                    Logger.WriteDebugLog("CyclewiseProductionDetails_LnT Report Exported successfully");
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("CyclewiseProductionDetails_LnT Report not mailed: no data");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            
        }
        #endregion

        public static void ExportMachinewiseShiftFormat1Report(DateTime StartDate, DateTime EndDate, string strReportFile, string ExportPath,string ExportFileName, string ShiftIn, string MachineID, string ComponentID, string OperationNo, string PlantID, bool Email_Flag, string Email_List_To,string Email_List_CC,string Email_List_BCC)
        {
            string dataAvailable = "NO";
            try
            {
                string dst = string.Empty;
                //string reportName = "SM_ShiftProductionReport.xlsx";

                if (!File.Exists(strReportFile))
                {
                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
                    return;
                }

                if (!Directory.Exists(ExportPath))
                {
                    Directory.CreateDirectory(ExportPath);
                }

                dst = Path.Combine(ExportPath, string.Format("SM_ShiftProductionReport_{0:ddMMMyyyyHHmmss}.xlsx", StartDate));

                if (File.Exists(dst))
                {
                    var di = new DirectoryInfo(dst);
                    di.Attributes &= ~FileAttributes.ReadOnly;
                    File.Delete(dst);
                }

                //string tempfileName = "SM_ShiftProductionReport" + "_" + Guid.NewGuid() + ".xlsx";
                //dst = Path.Combine(appPath, "Temp", SafeFileName(tempfileName));

                if (ShiftIn.Equals("All", StringComparison.OrdinalIgnoreCase))
                    ShiftIn = "";
                if (MachineID.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                    MachineID = "";
                if (PlantID.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                    PlantID = "";

                File.Copy(strReportFile, dst, true);
                FileInfo newFile = new FileInfo(dst);
                ExcelPackage pck = new ExcelPackage(newFile, true);
                var wsDts = pck.Workbook.Worksheets["Sheet1"];

                //string str1 = VDGDataBaseAccess.GetLogicalDayStart(StartDate.ToString("yyyy-MM-dd HH:mm:ss"));
                //DateTime strfromdate = DateTime.Now;
                //DateTime.TryParse(str1, out strfromdate);

                //string str2 = VDGDataBaseAccess.GetLogicalDayEnd(EndDate.ToString("yyyy-MM-dd HH:mm:ss"));
                //DateTime strtodate = DateTime.Now;
                //DateTime.TryParse(str2, out strtodate);

                wsDts.Cells["B3"].Value = StartDate.ToString("dd-MMM-yyyy");
                wsDts.Cells["E3"].Value = EndDate.ToString("dd-MMM-yyyy");
                int row = 0;
                System.Data.DataTable dt = AccessReportData.AnalysisMachinewiseShiftFormat1Report(StartDate, ShiftIn, MachineID, ComponentID, OperationNo, PlantID, EndDate, "Shift");
                if (dt != null && dt.Rows.Count > 0)
                {
                    dataAvailable = "YES";
                    row = 8;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        wsDts.Cells[row, 1].Value = dt.Rows[i]["shift"];
                        wsDts.Cells[row, 2].Value = Convert.ToDateTime(dt.Rows[i]["Day"]).ToString("dd/MM/yyyy");
                        wsDts.Cells[row, 3].Value = dt.Rows[i]["OperatorId"];
                        wsDts.Cells[row, 4].Value = dt.Rows[i]["MachineId"];
                        wsDts.Cells[row, 5].Value = dt.Rows[i]["MachineDescription"];
                        wsDts.Cells[row, 6].Value = dt.Rows[i]["Component"];
                        if (!string.IsNullOrEmpty(dt.Rows[i]["Operation"].ToString()))
                            wsDts.Cells[row, 7].Value = Convert.ToDecimal(dt.Rows[i]["Operation"]);
                        wsDts.Cells[row, 8].Value = dt.Rows[i]["OperationCount"];
                        wsDts.Cells[row, 9].Value = dt.Rows[i]["frmtCycleTime"];
                        wsDts.Cells[row, 10].Value = dt.Rows[i]["frmtAvgCycleTime"];
                        if (!(dt.Rows[i]["cyclefficiency"] is DBNull))
                            wsDts.Cells[row, 11].Value = Convert.ToDouble(dt.Rows[i]["cyclefficiency"]);
                        wsDts.Cells[row, 12].Value = dt.Rows[i]["frmtLoadUnload"];
                        wsDts.Cells[row, 13].Value = dt.Rows[i]["frmtAvgLoadUnload"];
                        if (!(dt.Rows[i]["LoadUnloadefficiency"] is DBNull))
                            wsDts.Cells[row, 14].Value = Convert.ToDouble(dt.Rows[i]["LoadUnloadefficiency"]);
                        wsDts.Cells[row, 15].Value = dt.Rows[i]["frmtUtilisedTime"];
                        wsDts.Cells[row, 16].Value = dt.Rows[i]["frmtDownTime"];
                        if (!(dt.Rows[i]["AvailabilityEfficiency"] is DBNull))
                            wsDts.Cells[row, 17].Value = Convert.ToDouble(dt.Rows[i]["AvailabilityEfficiency"]);
                        if (!(dt.Rows[i]["ProductionEfficiency"] is DBNull))
                            wsDts.Cells[row, 18].Value = Convert.ToDouble(dt.Rows[i]["ProductionEfficiency"]);
                        if (!(dt.Rows[i]["OverallEfficiency"] is DBNull))
                            wsDts.Cells[row, 19].Value = Convert.ToDouble(dt.Rows[i]["OverallEfficiency"]);//decimal.Round(, 2, MidpointRounding.AwayFromZero);
                        row++;
                    }
                    string modelRange = "A8:S" + (row - 1).ToString();
                    var modelTable = wsDts.Cells[modelRange];
                    //Added by abhi to hide columns
                    if (ConfigurationManager.AppSettings["AdvikPages"].ToString().Equals("1"))
                    {
                        wsDts.Column(11).Hidden = true;
                        wsDts.Column(14).Hidden = true;
                        wsDts.Column(17).Hidden = true;
                        wsDts.Column(18).Hidden = true;
                    }
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //DownloadMultipleFile(dst, pck.GetAsByteArray());
                    
                }

                pck.SaveAs(newFile);
                Logger.WriteDebugLog("Production report-Machinewise Shift-Format- 1 Report generated sucessfully.");

                if (dataAvailable.Equals("YES", StringComparison.OrdinalIgnoreCase))
                {
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportFileName);
                    Logger.WriteDebugLog("Production report-Machinewise Shift-Format- 1 Report Exported sucessfully.");
                }
                else
                {
                    Logger.WriteDebugLog("Production report-Machinewise Shift-Format- 1 Report not Mailed: No Data.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            
        }
    }
    public class MachineHistory
    {
        public string MachineID { get; set; }
        public string DownCode { get; set; }
        public string DownDescription { get; set; }
        public string KindOfProblem { get; set; }
        public string Reason { get; set; }
        public string DownCategory { get; set; }
        public string BreakDownStart { get; set; }
        public string BreakDownEnd { get; set; }
        public string ActionToResolve { get; set; }
        public string ActionProposed { get; set; }
        public string TimeLost { get; set; }
        public string ElapsedTime { get; set; }
        public string Severity { get; set; }
    }
}

#region EWS OEE Weekly Groupwise average plot
//internal static void ExportWeeklyEWSOEEReport(string strtTime, string endTime, string Shift, string strReportFile, string ExportPath, string ExportedReportFile,
//           string MachineId, string operators, string sttime,
//           string plantid, bool Email_Flag, string Email_List_To, string Email_List_CC,
//           string Email_List_BCC, int ShiftID)
//        {
//            string dst = string.Empty;
//            strtTime = "2018-06-04 06:00:00";
//            endTime = "2018-06-09 06:00:00";
//            try
//            {
//                bool dataAvailable = false;
//                if (!File.Exists(strReportFile))
//                {
//                    Logger.WriteDebugLog("Template is not found on " + strReportFile);
//                    return;
//                }

//                if (!Directory.Exists(ExportPath))
//                {
//                    Directory.CreateDirectory(ExportPath);
//                }

//                dst = Path.Combine(ExportPath, string.Format("SM_EWSWeeklyOEEReport_{0:ddMMMyyyy}.xlsx", DateTime.Parse(strtTime)));

//                if (File.Exists(dst))
//                {
//                    var di = new DirectoryInfo(dst);
//                    di.Attributes &= ~FileAttributes.ReadOnly;
//                    File.Delete(dst);
//                }
//                File.Copy(strReportFile, dst, true);

//                if (MachineId.Equals("ALL", StringComparison.OrdinalIgnoreCase))
//                {
//                    MachineId = "";
//                }

//                FileInfo newFile = new FileInfo(dst);
//                ExcelPackage excelPackage = new ExcelPackage(newFile, true);
//                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];
//                ws.Name = "Weekly OEE";

//                ws.Cells["C3"].Value = DateTime.Parse(strtTime).ToString("dd-MM-yy HH:mm:ss");
//                ws.Cells["F3"].Value = DateTime.Parse(endTime).ToString("dd-MM-yy HH:mm:ss");

//                System.Data.DataTable dt = AccessReportData.GetEWSWeeklyOEEData(strtTime, endTime, MachineId, plantid);
//                if (dt.Rows.Count > 0) dataAvailable = true;

//                int r = 7;
//                string prevGrp = "";
//                int prevRow = r;

//                DateTime stDate = DateTime.Parse(strtTime);
//                for (int i = 0; i < 6; i++)
//                {
//                    ws.Cells[r - 1, 4 + i].Value = stDate.AddDays(i).ToString("yyyy-MMM-dd");
//                }

//                Dictionary<string, int> rowList = new Dictionary<string, int>();
//                Dictionary<string, int> gantryList = new Dictionary<string, int>();
//                Dictionary<string, int> roboList = new Dictionary<string, int>();

//                foreach (DataRow row in dt.Rows)
//                {

//                    ws.Cells[r, 1].Value = row["Line"];
//                    ws.Cells[r, 2].Value = row["MachineID"];
//                    if (!prevGrp.Equals(row["Group"].ToString()))
//                    {
//                        ws.Cells[r, 3].Value = row["Group"];
//                        if (prevRow != r)
//                        {
//                            ws.Cells[prevRow, 3, r - 1, 3].Merge = true;
//                            ws.Cells[prevRow, 3, r - 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
//                            ws.Cells[prevRow, 11, r - 1, 11].Merge = true;
//                            ws.Cells[prevRow, 11, r - 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
//                            ws.Cells[prevRow, 11].Formula = string.Format("AVERAGE(J{0},J{1})", prevRow, r-1);
//                            rowList.Add(prevGrp, prevRow);
//                        }
//                        prevGrp = row["Group"].ToString();
//                        prevRow = r;
//                    }
//                    ws.Cells[r, 4].Value = row[3];
//                    ws.Cells[r, 5].Value = row[4];
//                    ws.Cells[r, 6].Value = row[5];
//                    ws.Cells[r, 7].Value = row[6];
//                    ws.Cells[r, 8].Value = row[7];
//                    ws.Cells[r, 9].Value = row[8];
//                    ws.Cells[r, 10].Value = float.Parse(row["OEE"].ToString());
//                    r++;
//                }

//                //if (!prevGrp.Equals("")) rowList.Add(prevGrp, prevRow);

//                if (prevRow != r)
//                {
//                    ws.Cells[prevRow, 3, r - 1, 3].Merge = true;
//                    ws.Cells[prevRow, 3, r - 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
//                    ws.Cells[prevRow, 11, r - 1, 11].Merge = true;
//                    ws.Cells[prevRow, 11, r - 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
//                    ws.Cells[prevRow, 11].Formula = string.Format("AVERAGE(J{0},J{1})", prevRow, r - 1);
//                    rowList.Add(prevGrp, prevRow);
//                }


//                ws.Cells[6, 1, r, 10].AutoFitColumns();

//                int pixeltop = 300;
//                int pixelleft = 1025;

//                for (int i = 0; i < rowList.Count - 1; i++)
//                {
//                    int l1 = rowList.ElementAt(i).Value;
//                    int l2 = rowList.ElementAt(i + 1).Value - 1;
//                    var chart = (ExcelBarChart)ws.Drawings.AddChart(string.Format("barChart{0}", i), eChartType.ColumnClustered);
//                    chart.SetSize(300, 300);
//                    chart.SetPosition(pixeltop * i + 20, pixelleft);
//                    chart.Title.Text = rowList.ElementAt(i).Key;
//                    chart.Series.Add(ExcelRange.GetAddress(l1, 10, l2, 10), ExcelRange.GetAddress(l1, 1, l2, 1));
//                }


//                int tmpr = rowList.ElementAt(rowList.Count - 1).Value;
//                var chart2 = (ExcelBarChart)ws.Drawings.AddChart(string.Format("barChart{0}", rowList.Count - 1), eChartType.ColumnClustered);
//                chart2.SetSize(300, 300);
//                chart2.SetPosition(pixeltop * (rowList.Count - 1) + 20, pixelleft);
//                chart2.Title.Text = rowList.ElementAt(rowList.Count - 1).Key;
//                chart2.Series.Add(ExcelRange.GetAddress(tmpr, 10, r - 1, 10), ExcelRange.GetAddress(tmpr, 1, r - 1, 1));

//                excelPackage.SaveAs(newFile);

//                if (dataAvailable)
//                {
//                    Logger.WriteDebugLog("EWS OEE Weekly Report Exported successfully");
//                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, dst, ExportedReportFile);
//                }
//                else
//                {
//                    Logger.WriteDebugLog("EWS OEE Weekly Report not mailed: no data");
//                }
//            }
//            catch (Exception ex)
//            {
//                Logger.WriteErrorLog("Error: " + ex.ToString());
//                Logger.WriteErrorLog(ex.StackTrace);
//            }
//        }
#endregion
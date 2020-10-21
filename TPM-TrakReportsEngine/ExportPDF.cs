using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Configuration;
using System.Reflection;

namespace TPM_TrakReportsEngine
{
    public static class ExportPDF
    {
        public static void createMCOreportPDF(string strReportFile, string ExportPath, string ExportedReportFile, int DayBefores, string frmdate, string todate, string plant, string machine, string shift, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {

            SqlConnection DBconection = ConnectionManager.GetConnection();
            DataTable dtColumnOrder1 = new DataTable();
            SqlDataReader sdr = null;
            SqlCommand Cmd = null;
            try
            {
                string path;
                Logger.WriteDebugLog(string.Format("Inconsistent MCO, report generating......"));

                Cmd = new SqlCommand("s_GetInconsistencyMCO", DBconection);
                Cmd.CommandType = CommandType.StoredProcedure;
                Cmd.CommandTimeout = 60 * 10;
                Cmd.Parameters.AddWithValue("@Fromdate", DateTime.Parse(frmdate).AddDays(DayBefores).ToString("yyyy-MM-dd HH:mm:ss"));
                Cmd.Parameters.AddWithValue("@Todate", DateTime.Parse(todate).AddDays(DayBefores).ToString("yyyy-MM-dd HH:mm:ss"));              
                if (plant.ToLower() == "all" || string.IsNullOrEmpty(plant))
                    Cmd.Parameters.AddWithValue("@PlantID", "");
                else
                {
                    Cmd.Parameters.AddWithValue("@PlantID", plant);
                }

                if (machine.ToLower() == "all" || string.IsNullOrEmpty(machine))
                {
                    Cmd.Parameters.AddWithValue("@Machineid", "");
                }
                else
                {
                    Cmd.Parameters.AddWithValue("@Machineid", machine);
                }
                sdr = Cmd.ExecuteReader();
                if (sdr.HasRows)
                {

                    if (!Directory.Exists(ExportPath))
                    {
                        Directory.CreateDirectory(ExportPath);
                    }
                    path = ExportPath + @ExportedReportFile + "_" + plant + "_" + string.Format("{0:ddMMMyyyy_HHmm}", DateTime.Parse(frmdate)) + ".pdf";
                    Document myDoc = new Document(PageSize.A4_LANDSCAPE.Rotate(), 10f, 10f, 20f, 30f);

                    string headerString = "From Date: " + frmdate + " To Date: " + todate;
                    headerString = headerString + " Plant: " + plant;
                    PlotData(dtColumnOrder1, headerString, sdr, myDoc, path, frmdate, todate, plant);
                    Logger.WriteDebugLog(ExportedReportFile + " Report generated sucessfully.");
                    sdr.Dispose();
                    DBconection.Close();
                    SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, path, ExportedReportFile);
                }
                else
                {
                    Logger.WriteDebugLog("No Inconsistent MCO record found for current shift/Day.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Report generation Failed. Error:{0}.", ex.ToString()));
                return;
            }
            finally
            {
                if(sdr != null) sdr.Dispose();
                if (DBconection != null) DBconection.Close();
            }
        }

        public static void PlotData(DataTable dtColumnOrder, string headerString, SqlDataReader sdr, Document myDoc, string path, string fromdate, string todate, string shift)
        {
            SqlConnection DBconection = null;
            try
            {
                DBconection = ConnectionManager.GetConnection();
                iTextSharp.text.Font fontHeaderTableH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 15f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font fontHeaderTableH2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                PdfPTable table = new PdfPTable(12);
                iTextSharp.text.Font fontH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f, iTextSharp.text.Font.BOLD, BaseColor.WHITE);
                iTextSharp.text.Font fontH2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f);

                # region Ploting Data in table

                PdfWriter pf = PdfWriter.GetInstance(myDoc, new FileStream(path, FileMode.Create));

                dtColumnOrder.Load(sdr);
                PdfPTable PdfTable = new PdfPTable(dtColumnOrder.Columns.Count);
                PdfPCell PdfPCell = null;

                PdfTable.TotalWidth = 100f;
                myDoc.Open();
                myDoc.Add(new Chunk("                                                                                                  "));
                myDoc.Add(new Chunk("Inconsistent data from machines", fontHeaderTableH1));
                myDoc.Add(new Paragraph(" "));
                myDoc.Add(new Chunk("                          "));
                myDoc.Add(new Chunk("From: " + string.Format("{0:dd-MMM-yyyy HH:mm:ss}", fromdate)));
                myDoc.Add(new Chunk("                          "));
                myDoc.Add(new Chunk("To: " + string.Format("{0:dd-MMM-yyyy HH:mm:ss}", todate)));
                myDoc.Add(new Chunk("                                      "));
                if (shift == string.Empty)
                {
                    myDoc.Add(new Chunk("Plant: " + "All"));
                }
                else
                {
                    myDoc.Add(new Chunk("Plant: " + shift));
                }


                if (dtColumnOrder != null)
                {
                    PdfPCell = new PdfPCell(new Phrase(new Chunk("MachineId", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("Machine", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("ComponentId", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("Component", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("Opn No.", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("Operator", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    PdfPCell = new PdfPCell(new Phrase(new Chunk("Remarks", fontH1)));
                    PdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    PdfPCell.BackgroundColor = new BaseColor(74, 154, 206);
                    PdfTable.AddCell(PdfPCell);

                    for (int rows = 0; rows < dtColumnOrder.Rows.Count; rows++)
                    {
                        for (int column = 0; column < dtColumnOrder.Columns.Count; column++)
                        {
                            PdfPCell = new PdfPCell(new Phrase(new Chunk(dtColumnOrder.Rows[rows][column].ToString(), fontH2)));
                            if (rows % 2 == 1)
                            {
                                PdfPCell.BackgroundColor = new BaseColor(197, 228, 226);
                            }
                            else
                            {
                                PdfPCell.BackgroundColor = new BaseColor(255, 255, 255);
                            }

                            PdfTable.AddCell(PdfPCell);
                        }
                    }
                }
                myDoc.Add(PdfTable);
                myDoc.Close();
                #endregion
            }
            catch (Exception ex)
            {
                if (myDoc != null) myDoc.Close();
                Logger.WriteErrorLog(string.Format("Report generation Failed. Error:{0}.", ex.ToString()));
                return;
            }
            finally
            {
                if (DBconection != null)
                    DBconection.Close();
            }
        }


        internal static void createCockpitreportPDF(string strReportFile, string exportFilePath,string exportFileName ,int DayBefores, string frmdate, string todate, string plant, string machine, string shift, string operatorVal, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {
            //createPDF("2015-03-03 06:00:00", "2015-03-04 06:00:00", "", "PCT", exportFilePath);
            createPDF(frmdate, todate, "", operatorVal, exportFilePath, exportFileName, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC);
        }


        public static void createPDF(string frmdate, string todate, string plantName, string user, string exportFilePath,string exportFileName, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC)
        {                       
            SqlConnection conection = null;
            try
            {
                conection = ConnectionManager.GetConnection();
                DataTable dtColumnOrder = new DataTable();
                SqlDataReader sdr = null;
                SqlCommand Cmd = null;

                // Cmd = new SqlCommand("select * from cockpitdefaults where parameter ='shopgridorder' order by ValueInInt", DBconection);
                Cmd = new SqlCommand("select valueintext,isnull(valueintext2,valueintext) as valueintext2,(select valueintext from userpreferences where parameter = cockpitdefaults.valueintext and EmployeeID='" + user + "' and modulename='SmartCockpit'  and formname='Frmcockpit' and controlname='MachineGrid') as valueintext from cockpitdefaults where parameter='shopgridorder' order by valueinint", conection);
             
                sdr = Cmd.ExecuteReader();
                dtColumnOrder.Load(sdr);
                conection.Close();

                Cmd = new SqlCommand("s_GetCockpitData", conection);
                Cmd.CommandType = CommandType.StoredProcedure;
                Cmd.Parameters.AddWithValue("@StartTime", frmdate);
                Cmd.Parameters.AddWithValue("@EndTime", todate);
                Cmd.Parameters.AddWithValue("@MachineID", "");
                Cmd.Parameters.AddWithValue("@PlantID", plantName);
               
                conection.Open();
                Cmd.CommandTimeout = 360;
                sdr = Cmd.ExecuteReader();
                if (!Directory.Exists(exportFilePath))
                {
                    Directory.CreateDirectory(exportFilePath);
                }
                string path = Path.Combine(exportFilePath, "Cockpit_" + string.Format("{0:ddMMMyyyy_HHmmss}", DateTime.Parse(frmdate)) + ".pdf");//TODO - add date time
                Document myDoc = new Document(PageSize.A4_LANDSCAPE.Rotate(), 20f, 20f, 20f, 30f);
                if (plantName == string.Empty) plantName = "All";
                PlotData(frmdate, todate, dtColumnOrder, sdr, myDoc, path, Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, exportFileName, plantName);
                sdr.Dispose();
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Error,.!!\n Creating PDF File..!!\n " + ex.Message);
            }
            finally
            {
                if (conection != null)
                {
                    conection.Close();
                }
            }
        }

        public static void PlotData(string fromDate,string toDate,DataTable dtColumnOrder, SqlDataReader sdr, Document myDoc, string path, bool Email_Flag, string Email_List_To, string Email_List_CC, string Email_List_BCC, string exportFileName,string plant)
        {         

            try
            {
                SqlConnection RevdbDBcon = ConnectionManager.GetConnection();
                SqlCommand RevCmd = null;
                SqlDataReader RevSdr = null;

                #region Header for the report
                iTextSharp.text.Font fontHeaderTableH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 15f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                PdfPTable headerTable = new PdfPTable(12);
                headerTable.TotalWidth = 800f;
                headerTable.LockedWidth = true;

                PdfPCell cellH1 = new PdfPCell(new Phrase("Machine wise Production Cockpit", fontHeaderTableH1));
                cellH1.Colspan = 12;
                cellH1.Border = 0;
                cellH1.HorizontalAlignment = 1;
                headerTable.AddCell(cellH1);
                #endregion

                PdfPTable table = new PdfPTable(12);
                iTextSharp.text.Font fontH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f, iTextSharp.text.Font.BOLD, BaseColor.WHITE);
                iTextSharp.text.Font fontH2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f);
                table.TotalWidth = 800f;
                table.LockedWidth = true;

                # region Header for the table

                /* Prints the Header of the Cockpit table 
                ----------------------------------------------------*/
                PdfPHeaderCell hcell1 = new PdfPHeaderCell();
                hcell1.Column.AddText(new Phrase("MachineID", fontH1));
                hcell1.BackgroundColor = iTextSharp.text.BaseColor.GRAY;
                hcell1.HorizontalAlignment = Element.ALIGN_CENTER;
                table.AddCell(hcell1);

                for (int i = 0; i < dtColumnOrder.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(dtColumnOrder.Rows[i]["ValueInText2"].ToString()))
                    {
                        PdfPHeaderCell hcell2 = new PdfPHeaderCell();
                        hcell2.Column.AddText(new Phrase(dtColumnOrder.Rows[i]["ValueInText2"].ToString(), fontH1));
                        hcell2.BackgroundColor = iTextSharp.text.BaseColor.GRAY;
                        hcell2.HorizontalAlignment = Element.ALIGN_CENTER;
                        table.AddCell(hcell2);
                    }

                    else
                    {
                        PdfPHeaderCell hcell2 = new PdfPHeaderCell();
                        string hValue = dtColumnOrder.Rows[i]["ValueInText"].ToString();
                        hValue = hValue.Replace("E", " E");
                        hValue = hValue.Replace("L", " L");
                        hValue = hValue.Replace("T", " T");
                        hcell2.Column.AddText(new Phrase(hValue, fontH1));
                        // hcell2.Column.AddText(new Phrase(dtColumnOrder.Rows[i]["ValueInText"].ToString(), fontH1));
                        hcell2.BackgroundColor = iTextSharp.text.BaseColor.GRAY;
                        hcell2.HorizontalAlignment = Element.ALIGN_CENTER;
                        table.AddCell(hcell2);
                    }
                }

                /*End of Printing Header column
                 ----------------------------------------------------*/
                #endregion

                # region Ploting Data in table

                if (sdr.HasRows)
                {
                    int j = 0;
                    try
                    {
                        while (sdr.Read())
                        {
                            PdfPCell cell1 = new PdfPCell(new Phrase(sdr["MachineID"].ToString(), fontH2));
                            if (j % 2 == 1)
                            {
                                cell1.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                            }
                            else
                            {
                                cell1.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                            }
                            table.AddCell(cell1);

                            PdfPCell cell2 = new PdfPCell();

                            for (int k = 0; k < dtColumnOrder.Rows.Count; k++)
                            {
                                if (!string.IsNullOrEmpty(dtColumnOrder.Rows[k]["ValueInText"].ToString()))
                                {
                                    string columName = dtColumnOrder.Rows[k]["ValueInText"].ToString();

                                    #region Based on column name

                                    if (columName == "AvailabilityEfficiency")
                                    {
                                        /* Color ploting for AvailabilityEfficiency" */
                                        double AE = Convert.ToDouble(sdr["AvailabilityEfficiency"].ToString());
                                        AE = Math.Round(AE, 2);
                                        cell2 = new PdfPCell(new Phrase(Convert.ToString(AE), fontH2));

                                        if (Convert.ToDouble(AE) <= Convert.ToDouble(sdr["AERed"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.RED;
                                        }
                                        else if (Convert.ToDouble(AE) > Convert.ToDouble(sdr["AERed"]) && Convert.ToDouble(AE) < Convert.ToDouble(sdr["AEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                                        }
                                        else if (Convert.ToDouble(AE) >= Convert.ToDouble(sdr["AEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.GREEN;
                                        }
                                    }

                                    if (columName == "P.Efficiency")
                                    {
                                        double PE = Convert.ToDouble(sdr["ProductionEfficiency"].ToString());
                                        PE = Math.Round(PE, 2);
                                        cell2 = new PdfPCell(new Phrase(Convert.ToString(PE), fontH2));

                                        if (Convert.ToDouble(PE) <= Convert.ToDouble(sdr["PERed"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.RED;
                                        }
                                        else if (Convert.ToDouble(PE) > Convert.ToDouble(sdr["PERed"]) && Convert.ToDouble(PE) < Convert.ToDouble(sdr["PEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                                        }
                                        else if (Convert.ToDouble(PE) >= Convert.ToDouble(sdr["PEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.GREEN;
                                        }
                                    }

                                    if (columName == "OverAllEfficiency")
                                    {
                                        /* Color ploting for OverAllEfficiency" */
                                        double OE = Convert.ToDouble(sdr["OverAllEfficiency"].ToString());
                                        OE = Math.Round(OE, 2);
                                        cell2 = new PdfPCell(new Phrase(Convert.ToString(OE), fontH2));

                                        if (Convert.ToDouble(OE) <= Convert.ToDouble(sdr["OERed"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.RED;
                                        }
                                        else if (Convert.ToDouble(OE) > Convert.ToDouble(sdr["OERed"]) && Convert.ToDouble(OE) < Convert.ToDouble(sdr["OEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                                        }
                                        else if (Convert.ToDouble(OE) >= Convert.ToDouble(sdr["OEGreen"]))
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.GREEN;
                                        }
                                    }

                                    if (columName == "Components")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["Components"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "UtilisedTime")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["StrUtilisedTime"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "ManagementLoss")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["DownTime"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "DownTime")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["DownTime"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "TotalTime")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["TotalTime"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "MaxReasonTime")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["MaxReasonTime"].ToString(), fontH2));

                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "Remarks")
                                    {
                                        cell2 = new PdfPCell(new Phrase(sdr["Remarks"].ToString(), fontH2));
                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    if (columName == "RevenueLoss")
                                    {
                                        double revValue = 0;
                                        string tmtformat = string.Empty;
                                        string dwntime = string.Empty;
                                        double RV = 0;

                                        dwntime = sdr["DownTime"].ToString();

                                        RevCmd = new SqlCommand("select mchrrate from machineinformation where machineid= '" + sdr["MachineID"].ToString() + "' ", RevdbDBcon);
                                        RevdbDBcon = ConnectionManager.GetConnection();
                                        SqlDataReader RevSdr1 = RevCmd.ExecuteReader();
                                        if (RevSdr1.HasRows)
                                        {
                                            RevSdr1.Read();
                                            revValue = Convert.ToDouble(RevSdr1[0]);
                                        }
                                        RevdbDBcon = ConnectionManager.GetConnection();

                                        RevCmd = new SqlCommand("select valueinText from cockpitdefaults where Parameter='TimeFormat'", RevdbDBcon);
                                        RevdbDBcon = ConnectionManager.GetConnection();
                                        SqlDataReader RevSdr2 = RevCmd.ExecuteReader();
                                        if (RevSdr2.HasRows)
                                        {
                                            RevSdr2.Read();
                                            tmtformat = RevSdr2[0].ToString();
                                        }
                                        RevdbDBcon = ConnectionManager.GetConnection();
                                        RV = Moneyval(dwntime, tmtformat, revValue);

                                        cell2 = new PdfPCell(new Phrase(RV.ToString(), fontH2));
                                        if (j % 2 == 1)
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.CYAN;
                                        }
                                        else
                                        {
                                            cell2.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
                                        }
                                    }

                                    #endregion

                                    table.AddCell(cell2);
                                }
                            }
                            j++;
                        }
                    }
                    catch (Exception ex)
                    {

                        Logger.WriteErrorLog(ex.ToString());
                    }
                }
                #endregion
                PdfWriter pdfWriter = null;
                try
                {
                    pdfWriter = PdfWriter.GetInstance(myDoc, new FileStream(path, FileMode.OpenOrCreate));
                }
                catch (Exception exxx)
                {
                    path = path.Replace(".pdf", Path.GetRandomFileName() + ".pdf");
                    pdfWriter = PdfWriter.GetInstance(myDoc, new FileStream(path, FileMode.OpenOrCreate));
                }
                myDoc.NewPage();
                myDoc.Open();
                myDoc.Add(headerTable);
                myDoc.Add(new Paragraph(" "));
                myDoc.Add(new Chunk("From: " + string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(fromDate))));
                myDoc.Add(new Chunk("                                      "));
                myDoc.Add(new Chunk("      "));
                myDoc.Add(new Chunk("To: " + string.Format("{0:dd-MMM-yyyy}", DateTime.Parse(toDate))));
                myDoc.Add(new Chunk("                                      "));
                myDoc.Add(new Chunk("                              "));
                myDoc.Add(new Chunk("Plant: " + plant));
                myDoc.Add(new Paragraph(" "));
                myDoc.Add(table);
                myDoc.Close();
                pdfWriter.Close();
                SendEmail.SendEmailMsg(Email_Flag, Email_List_To, Email_List_CC, Email_List_BCC, path, exportFileName);
              

                #region CommentedCode
                //myDoc.NewPage();            
                //myDoc.Open();
                //myDoc.Add(new Chunk("Page 2"));

                //myDoc.Open();
                //headerTable.WriteSelectedRows(0, -1, myDoc.Left, myDoc.Top, pf.DirectContent);
                //table.WriteSelectedRows(0, -1, 20f, 500f, pf.DirectContent); 
                #endregion

            }
            catch (Exception exx)
            {
                Logger.WriteDebugLog("Error,.!!\n Creating PDF File..!!\n " + exx.Message);
            }
        }

        public static double Moneyval(string DwnTime, string tmformat, double macval)
        {
            string[] iparray = new string[10];
            long tmsec;
            double tmhour;
            double RevValue = 0;

            try
            {
                tmsec = 0;

                if (tmformat == "hh:mm:ss")
                {
                    iparray = DwnTime.Split(':');
                    tmsec = Convert.ToInt32(iparray[0]) * 3600;
                    tmsec = tmsec + (Convert.ToInt32(iparray[1]) * 60);
                    tmsec = tmsec + Convert.ToInt32(iparray[2]);
                }
                else if (tmformat == "hh")
                {
                    iparray = DwnTime.Split('.');
                    tmsec = Convert.ToInt32(iparray[0]) * 3600;
                    tmsec = tmsec + Convert.ToInt32(iparray[1]) * 60;
                }
                else if (tmformat == "mm")
                {
                    iparray = DwnTime.Split('.');
                    tmsec = Convert.ToInt32(iparray[0]) * 60;
                    tmsec = tmsec + Convert.ToInt32(iparray[1]);
                }
                else
                {
                    tmsec = Convert.ToInt32(DwnTime);
                }

                if (tmsec != 0)
                {
                    tmhour = tmsec / 3600;
                }
                else
                {
                    tmhour = 0;
                }

                RevValue = tmhour * macval;
                RevValue = Math.Round(RevValue, 2);
                return (RevValue);
            }
            catch
            {
                return 0.0;
            }
        }
    }
}

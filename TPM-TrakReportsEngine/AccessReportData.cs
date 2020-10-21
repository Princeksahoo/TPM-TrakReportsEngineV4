using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace TPM_TrakReportsEngine
{
    class AccessReportData
    {
        public static string APath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public static SqlDataReader GetRunreportType()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Select * from scheduledreports", Con);
            SqlDataReader reader = null;
            try
            {
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            return reader;
        }

        public static List<string> GetExportReportPaths()
        {
            DateTime shiftEndTime = DateTime.MinValue;
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT DISTINCT ExportPath FROM ScheduledReports", Con);
            cmd.CommandTimeout = 60;
            List<string> paths = new List<string>();
            SqlDataReader reader = null;
            DataTable dt = new DataTable();
            try
            {
                reader = cmd.ExecuteReader();
                dt.Load(reader);
                foreach (DataRow row in dt.Rows)
                {
                    paths.Add(row["ExportPath"].ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            return paths;
        }

        public static SqlDataReader GetExportReports(string Vartime)
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            string Query = string.Empty;

            if (string.IsNullOrEmpty(Vartime))
            {
                //  Query = "SELECT * FROM ScheduledReports where RunReportForEvery <> 'Now' and RunReportForEvery <>'Month' and RunReportForEvery <>'Weekly' order by Slno";

                Query = @"select Top 1 * From ScheduledReports where ISNULL(ReportID,0) ='24' and RunReportForEvery <>'Month' and RunReportForEvery <>'Weekly'
                UNION
                select top 1 * from ScheduledReports where ISNULL(ReportID,0) ='25' and RunReportForEvery <>'Month' and RunReportForEvery <>'Weekly'
                union
                Select * From ScheduledReports where ISNULL(ReportID,0) not in('24','25') and RunReportForEvery <>'Month' and RunReportForEvery <>'Weekly'
                Order by slno";

            }
            else if (Vartime.ToLower() == "now")
            {
                Query = "SELECT * FROM ScheduledReports where runreportforevery = '" + Vartime + "' and RunHistory is null order by Slno";
            }

            else
            {
                Query = "SELECT * FROM ScheduledReports where runreportforevery = '" + Vartime + "' order by Slno";
            }

            SqlCommand cmd = new SqlCommand(Query, Con);
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static int MaxHourIdShift()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Select max(hourId)as Col from shifthourdefinition", Con);

            object MaxHourId = null;
            try
            {
                MaxHourId = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }

            return (MaxHourId == DBNull.Value || MaxHourId == null) ? 0 : int.Parse(MaxHourId.ToString());
        }

        public static SqlDataReader ProdDownReport(DateTime StartDate, DateTime EndDate, string PlantID, string MachineID, string RptProd_down)
        {

            SqlConnection Con = ConnectionManager.GetConnection();

            SqlCommand cmd = new SqlCommand("s_Get_Auma_Prod_Downreport", Con);
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate;
            cmd.Parameters.Add("@Enddate", SqlDbType.DateTime).Value = EndDate;
            cmd.Parameters.Add("@PlantID", SqlDbType.NVarChar).Value = PlantID;
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@RptProd_down", SqlDbType.NVarChar).Value = RptProd_down;
            cmd.CommandTimeout = 360;
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static SqlDataReader DNCUsageReport(DateTime StartDate, DateTime EndDate, string MachineID)
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("s_GetDNCUsage", Con);
            cmd.Parameters.Add("@starttime", SqlDbType.DateTime).Value = StartDate;
            cmd.Parameters.Add("@endtime", SqlDbType.DateTime).Value = EndDate;
            cmd.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = string.Empty;
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@clientname", SqlDbType.NVarChar).Value = string.Empty;
            cmd.Parameters.Add("@Param", SqlDbType.NVarChar).Value = "Successful Transfer with QTY";
            cmd.CommandTimeout = 360;
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static SqlDataReader ShiftProductionCountHour(DateTime StartDate, string MachineID, string Param)
        {

            SqlConnection Con = ConnectionManager.GetConnection();

            SqlCommand cmd = new SqlCommand("s_GetHourlyTarget_Count_followup", Con);
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate;
            cmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = StartDate;
            cmd.Parameters.Add("@Shift", SqlDbType.NVarChar).Value = string.Empty;
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@Param", SqlDbType.NVarChar).Value = Param;
            cmd.CommandTimeout = 360;
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static string GetCompanyName()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select top 1 companyname from company", Con);

            object CompanyName = null;
            try
            {
                CompanyName = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (CompanyName == null || Convert.IsDBNull(CompanyName))
            {
                return string.Empty;
            }
            return (string)CompanyName;
        }

        public static bool GetPDT(string LRunDay)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select * from HolidayList where Holiday = '" + string.Format("{0:yyyy-MMM-dd}", DateTime.Parse(LRunDay)) + "'", Con);

            object PDT = null;
            try
            {
                PDT = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return (PDT == DBNull.Value || PDT == null) ? false : true;
        }

        public static SqlDataReader GetExportReports(string Vartime, int ShiftId)
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            string Query = string.Empty;

            if (ShiftId == 3)
            {
                Query = "SELECT * FROM ScheduledReports order by Slno";
            }
            else
            {
                Query = "SELECT * FROM ScheduledReports where runreportforevery = '" + Vartime + "' order by Slno";
            }

            SqlCommand cmd = new SqlCommand(Query, Con);
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        #region ShiftReleatedFunctions
        //public static int GetShiftID()
        //{
        //    SqlConnection Con = ConnectionManager.GetConnection();
        //    string TimeSec = string.Format("{0:HH:mm:ss}", DateTime.Now);
        //    SqlCommand cmd = new SqlCommand("select top 1 ShiftID from shiftdetails where running=1 and convert(datetime,CAST(datePart(hh,ToTime) AS nvarchar(2)) + ':' + CAST(datePart(mi,ToTime) as nvarchar(2))+ ':' + CAST(datePart(ss,ToTime) as nvarchar(2)))<'" + string.Format("{0:HH:mm:ss}", DateTime.Now) + "' order by totime desc", Con);
        //    object ShiftID = null;
        //    try
        //    {
        //        ShiftID = cmd.ExecuteScalar();
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.WriteErrorLog(ex);
        //    }
        //    finally
        //    {
        //        if (Con != null)
        //        {
        //            Con.Close();
        //        }
        //    }

        //    return (ShiftID == DBNull.Value || ShiftID == null) ? 0 : int.Parse(ShiftID.ToString());
        //}

        public static SqlDataReader GetPreviousShiftEndTime() //Todo
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            //DR0331::Geeta Added from here
            SqlCommand cmd = new SqlCommand("s_GetPreviousShift", Con);  /* returns only current shift Start-End Time */
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;

            //DR0331::Geeta Added till here
            //DR0331::Geeta Commented from here

            //SqlCommand cmd = new SqlCommand("select top 1 Fromtime,Totime,Today, Fromday, ShiftName,ShiftID from shiftdetails where running=1 and convert(datetime,CAST(datePart(hh,ToTime) AS nvarchar(2)) + ':' + CAST(datePart(mi,ToTime) as nvarchar(2))+ ':' + CAST(datePart(ss,ToTime) as nvarchar(2)))<'" + string.Format("{0:HH:mm:ss}", DateTime.Now) + "' order by totime DESC", Con);
            //SqlDataReader dr = null;
            //try
            //{
            //    dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            //}
            //catch (Exception ex)
            //{
            //    Logger.WriteErrorLog(ex);
            //}
            //return dr;
            //DR0331::Geeta Commented till here
        }

        public static int GetShiftIDNO(string SWName)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            string TimeSec = string.Format("{0:HH:mm:ss}", DateTime.Now);
            SqlCommand cmd = new SqlCommand("select top 1 ShiftID from shiftdetails where running = 1 and ShiftName = '" + SWName + "'", Con);
            object ShiftID = null;
            try
            {
                ShiftID = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return (ShiftID == DBNull.Value || ShiftID == null) ? 0 : int.Parse(ShiftID.ToString());
        }

        public static SqlDataReader GetCurrentShiftDetails()
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("s_GetCurrentShiftTime", Con);  /* returns only current shift Start-End Time */
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = DateTime.Now;
            cmd.Parameters.Add("@Param", SqlDbType.NVarChar).Value = "";
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static void GetCurrentShiftTimeDS(DateTime ShiftDT, out DateTime FromDT, out DateTime ToDT, out string shiftName)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("S_GetShiftTimeSA", Con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@StartDateTime", SqlDbType.DateTime).Value = ShiftDT;
            SqlDataAdapter DA = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            DA.Fill(ds, "ShiftTime");
            try
            {
                foreach (DataRow DR in ds.Tables["ShiftTime"].Rows)
                {
                    DateTime FFromTime = DateTime.Parse(DR["StartTime"].ToString());
                    DateTime FToTime = DateTime.Parse(DR["EndTime"].ToString());
                    string SName = Convert.ToString(DR["ShiftName"]);
                    DateTime CurTime = ShiftDT;

                    if (CurTime > FFromTime && CurTime < FToTime)
                    {
                        FromDT = FFromTime;
                        ToDT = CurTime;
                        shiftName = SName;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                    ds.Dispose();
                    DA.Dispose();
                    cmd.Dispose();
                }
            }
            FromDT = DateTime.Now;
            ToDT = DateTime.Now;
            shiftName = string.Empty;
        }

        public static string GetLogicalDay(string LRunDay, string Param)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDay( '" + string.Format("{0:yyyy-MMM-dd}", DateTime.Parse(LRunDay)) + "','" + Param + "')", Con);

            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MMM-dd HH:mm:ss tt}", Convert.ToDateTime(SEDate));
        }
        #endregion

        #region ShopDefaults

        public static string GetMachineAE()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select valueintext from shopdefaults where parameter='Machine AE' and valueintext = 'Time Consolidated' and isnull(valueintext,'')<>''", Con);

            object MachineAE = null;
            try
            {
                MachineAE = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (MachineAE == null || Convert.IsDBNull(MachineAE))
            {
                return string.Empty;
            }
            return (string)MachineAE;
        }

        public static DateTime GetLastRunforTheDay()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Select ValueInText from ShopDefaults where Parameter = 'ScheduledReports_LastRunforTheDay'", Con);

            object LastRunforTheDay = null;
            try
            {
                LastRunforTheDay = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (LastRunforTheDay == null || Convert.IsDBNull(LastRunforTheDay))
            {
                return DateTime.Now;
            }
            return DateTime.Parse((string)LastRunforTheDay);
        }

        public static SqlDataReader GetSendEmail()
        {

            SqlConnection Con = ConnectionManager.GetConnection();

            SqlCommand cmd = new SqlCommand("Select ValueinText,ValueinText2,valueinint from ShopDefaults where Parameter ='ScheduledReports_Email'", Con);
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static SqlDataReader GetMailServerDomain()
        {
            START:

            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Select * from ShopDefaults where Parameter ='ScheduledReports_MailServerDomain'", Con);

            SqlDataReader sdr = null;
            try
            {
                sdr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            if (sdr == null) goto START;
            return sdr;
        }

        public static int GetOverWriteFile()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select valueinInt from shopdefaults where parameter='OverWriteFile'", Con);

            object OverWriteFile = null;
            try
            {
                OverWriteFile = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }

            return (OverWriteFile == DBNull.Value || OverWriteFile == null) ? 0 : int.Parse(OverWriteFile.ToString());
        }

        public static SqlDataReader GetMailSubjectAndBody()
        {

            SqlConnection Con = ConnectionManager.GetConnection();

            SqlCommand cmd = new SqlCommand("select * from shopdefaults where parameter='ScheduledReportsEmail_Text'", Con);
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static void InsertLRunDay(string LRunDay)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Insert into ShopDefaults (Parameter,ValueInText) values('ScheduledReports_LastRunforTheDay','" + string.Format("{0:yyyy-MMM-dd HH:mm:ss}", DateTime.Parse(LRunDay)) + "'", Con);

            object PDT = null;
            try
            {
                PDT = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
        }

        public static void UpdateLRunDay(string LRunDay)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Update ShopDefaults set ValueInText = '" + string.Format("{0:yyyy-MMM-dd HH:mm:ss}", DateTime.Parse(LRunDay)) + "' where parameter = 'ScheduledReports_LastRunforTheDay'", Con);
            //SqlCommand cmd = new SqlCommand("Update ShopDefaults set ValueInText = '" + string.Format("{0:yyyy-MMM-dd}", DateTime.Parse(LRunDay)) + "' where parameter = 'ScheduledReports_LastRunforTheDay'", Con);

            object PDT = null;
            try
            {
                PDT = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
        }
        #endregion

        public static SqlDataReader OEETrend(string StTime, string EndTime, string Shift, string Plant, string Machine, string Param, string Format)
        {

            bool isInError = false;

            SqlConnection con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("s_GetOEETrend", con);
            cmd.Parameters.AddWithValue("@StartDate", SqlDbType.DateTime).Value = StTime.Trim();
            cmd.Parameters.AddWithValue("@EndDate", SqlDbType.DateTime).Value = EndTime.Trim();
            cmd.Parameters.AddWithValue("@shift", SqlDbType.NVarChar).Value = Shift.Trim();
            cmd.Parameters.AddWithValue("@PlantID", SqlDbType.NVarChar).Value = Plant.Trim();
            cmd.Parameters.AddWithValue("@MachineID", SqlDbType.NVarChar).Value = Machine;
            cmd.Parameters.AddWithValue("@Parameter", SqlDbType.NVarChar).Value = Param.Trim();
            cmd.Parameters.AddWithValue("@Format", SqlDbType.NVarChar).Value = Format.Trim();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 360;
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                isInError = true;
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }

        public static SqlDataReader ProductionTrelBorg(DateTime StartDate, DateTime EndDate, string PlantID, string MachineID, string componentId, string operationNo, int daysBefore)
        {

            bool isInError = false;

            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("s_GetDailyProdandDownReport_RuntimeByMCO", Con);
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate.AddDays(daysBefore);
            cmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate.AddDays(daysBefore);
            cmd.Parameters.Add("@PlantID", SqlDbType.NVarChar).Value = PlantID;
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@ComponentID", SqlDbType.NVarChar).Value = componentId;
            cmd.Parameters.Add("@OperationNo", SqlDbType.NVarChar).Value = operationNo;
            cmd.CommandTimeout = 60 * 60;
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult);
            }
            catch (Exception ex)
            {
                isInError = true;
                Logger.WriteErrorLog(ex);
            }

            return dr;
        }


        public static void UpdateScheduleReportTrelBorg(string seriolNo)
        {
            string sqlQuery = "update ScheduledReports set RunHistory=getdate() where Slno='" + seriolNo + "'";
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(sqlQuery, Con);
            try
            {
                int ret = cmd.ExecuteNonQuery();
                if (ret > 0)
                {
                    Logger.WriteDebugLog("ScheduleRerports table updated successfully for trelborg for serial no= " + seriolNo + ".");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                    Con.Close();
            }

        }

        public static string GetLogicalMonthStartEnd(DateTime currentDate, string Param)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalMonth( '" + string.Format("{0:yyyy-MM-dd}", currentDate) + "','" + Param + "')", Con);

            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }

        public static void UpdateScheduleDncReportMonthWise(string seriolNo)
        {
            string sqlQuery = "update ScheduledReports set RunHistory='" + DateTime.Now.ToString("yyyy-MMM-dd HH:mm:ss") + "' where Slno='" + seriolNo + "'";
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(sqlQuery, Con);
            try
            {
                int ret = cmd.ExecuteNonQuery();
                if (ret > 0)
                {
                    Logger.WriteDebugLog("ScheduleRerports table updated successfully for  serial no= " + seriolNo + ".");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                    Con.Close();
            }
        }

        public static List<string> GetTPMTrakEnabledMachines(string plantname)
        {
            if (plantname.Equals("All Plant", StringComparison.OrdinalIgnoreCase))
            {
                plantname = "";
            }
            List<string> machines = new List<string>();
            string sqlQuery = @"select PM.MachineID  from plantinformation PT inner  join plantmachine PM on 
            PT.PlantID = PM.PlantID
            inner join machineinformation MI on PM.MachineID= MI.machineid 
            where  (PT.Plantid=@plantname or @plantname ='') and  MI.TPMTrakEnabled=1";
            SqlDataReader reader = default(SqlDataReader);
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(sqlQuery, Con);
            cmd.Parameters.AddWithValue("@plantname", plantname);
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    machines.Add(reader.GetString(0));
                }

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (Con != null)
                    Con.Close();
            }
            return machines;
        }

        public static SqlDataReader GetMoReport(DateTime StartDate, DateTime EndDate, string PlantID, string MachineID, string param)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("S_GetMODetails", Con);
            cmd.Parameters.AddWithValue("@Starttime", StartDate.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@Endtime", EndDate.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.Add("@PlantID", SqlDbType.NVarChar).Value = PlantID;
            cmd.Parameters.Add("@Machineid", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@Param", SqlDbType.NVarChar).Value = param;
            cmd.CommandTimeout = 60 * 30;
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                Logger.WriteDebugLog("started exec proc S_GetMODetails");
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                Logger.WriteDebugLog("completed proc S_GetMODetails");
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            return dr;
        }

        public static string GetLogicalDayEnd(string LRunDay)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDayEnd( '" + string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Parse(LRunDay).AddSeconds(1)) + "')", Con);
            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }

        public static string GetLogicalDayStart(string LRunDay)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDayStart( '" + string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Parse(LRunDay).AddSeconds(1)) + "')", Con);

            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }
        //vasavi
        private void PlotGraphs(ExcelWorksheet wsDt, int startPos, int series, int EnergyColIndex)
        {
            var chart = (ExcelBarChart)wsDt.Drawings.AddChart("Energy Report", eChartType.ColumnClustered);
            chart.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;

            chart.SetSize(710, 310);
            chart.Title.Text = "Energy Report";
            chart.Legend.Remove();

            var serie1 = chart.Series.Add(ExcelRange.GetAddress(startPos, EnergyColIndex, series, EnergyColIndex), ExcelRange.GetAddress(startPos, 1, series, 1));
            chart.YAxis.Title.Text = "KWh";

            var chartz = (ExcelPieChart)wsDt.Drawings.AddChart("Energy Reportz", eChartType.Pie);
            if (EnergyColIndex <= 8)
            {
                chart.SetPosition((series + 13) * 20, 22);
                chartz.SetPosition(((series + 13) * 20), 766);
            }
            else
            {
                chart.SetPosition((series + 12) * 18, 22);
                chartz.SetPosition(((series + 12) * 18), 766);
            }
            chartz.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;

            chartz.SetSize(420, 310);
            chartz.Title.Text = "Energy Report";

            var serie2 = chartz.Series.Add(ExcelRange.GetAddress(startPos, EnergyColIndex, series, EnergyColIndex), ExcelRange.GetAddress(startPos, 1, series, 1));
            chartz.Legend.Position = eLegendPosition.Top;
        }


        internal static Dictionary<string, List<string>> GetCatAndSubCat()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select Category, SubCategory from pm_information group by Category, SubCategory", Con);
            SqlDataReader rdr;
            Dictionary<string, List<string>> dct = new Dictionary<string, List<string>>();

            try
            {
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string k = rdr["Category"].ToString();
                    string v = rdr["SubCategory"].ToString();
                    if (dct.ContainsKey(k))
                    {
                        dct[k].Add(v);
                    }
                    else
                    {
                        dct.Add(k, new List<string>(new[] { v }));
                    }
                }
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return dct;
        }

        public static DataTable GetShiftIDsandNames()
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable dt = new DataTable();
            try
            {
                SqlCommand cmd = new SqlCommand(@"SELECT ShiftName, ShiftID FROM ShiftDetails WHERE Running=1 ORDER BY ShiftID");
                cmd.CommandType = CommandType.Text;
                cmd.CommandTimeout = 600;
                cmd.Connection = sqlConn;
                SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(dr);
                dr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("GetShiftNames: " + ex.StackTrace);
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
            return dt;
        }

        public static SqlDataReader GetHourlyMachinewiseProduction(string strtTime, string endTime, string PlantID, string MachineID)
        {

            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetNSPL_Reports]", sqlConn);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 600;
            cmd.Parameters.AddWithValue("@StartDate", strtTime);
            cmd.Parameters.AddWithValue("@EndDate", endTime);
            cmd.Parameters.AddWithValue("@Machine", MachineID);
            cmd.Parameters.AddWithValue("@PlantID", PlantID);
            cmd.Parameters.AddWithValue("@ComparisonParam", "Shift");
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            return dr;
        }



        internal static SqlDataReader GetProductionEfficiency(string strtTime, string endTime, string plantid, string MachineId, string comparisonParam, string timeAxis, string shiftName, string type)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetEfficiencyFromAutodata]", sqlConn);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 600;
            cmd.Parameters.AddWithValue("@StartTime", strtTime);
            cmd.Parameters.AddWithValue("@EndTime", endTime);
            cmd.Parameters.AddWithValue("@MachineID", MachineId);
            cmd.Parameters.AddWithValue("@PlantID", plantid);
            cmd.Parameters.AddWithValue("@ComparisonParam", comparisonParam);
            cmd.Parameters.AddWithValue("@TimeAxis", timeAxis);
            cmd.Parameters.AddWithValue("@ShiftName", shiftName);
            cmd.Parameters.AddWithValue("@Type", type);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            return dr;
        }

        internal static SqlDataReader GetMandoReport(string strtTime, string endTime, string plantid, string MachineId, string shiftID, string sheetNo, string format)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetMando_Reports]", sqlConn);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 600;
            cmd.Parameters.AddWithValue("@StartDate", strtTime);
            cmd.Parameters.AddWithValue("@EndDate", endTime);
            cmd.Parameters.AddWithValue("@Machine", MachineId);
            cmd.Parameters.AddWithValue("@Plant", plantid);
            cmd.Parameters.AddWithValue("@ShiftID", shiftID);
            cmd.Parameters.AddWithValue("@SheetNo", sheetNo);
            cmd.Parameters.AddWithValue("@Format", format);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = null;
            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            return dr;
        }

        internal static DataTable GetPMReport(string startTime, string endTime, string machineId)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable dt = new DataTable();
            try
            {
                SqlCommand cmd = new SqlCommand(@"[s_GenerateShantiPMReport]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 1800;
                cmd.Parameters.AddWithValue("@Starttime", startTime);
                cmd.Parameters.AddWithValue("@Endtime", endTime);
                cmd.Parameters.AddWithValue("@MachineID", machineId.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : machineId);
                SqlDataReader rdr = cmd.ExecuteReader();
                dt.Load(rdr);
                rdr.Close();
            }
            catch (Exception e)
            {
                Logger.WriteErrorLog("GetPMReport: " + e.ToString());
            }
            return dt;
        }

        internal static Dictionary<string, List<string>> GetCatAndSubCat(string strMacName)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"
IF NOT EXISTS(select Category, SubCategory from PM_Information where MachineType=(SELECT TOP 1 Description FROM machineinformation WHERE  machineid=@machid))
SELECT Category, SubCategory from PM_Information where MachineType='GENERAL'
ELSE
SELECT Category, SubCategory from PM_Information where MachineType=(SELECT TOP 1 Description FROM machineinformation WHERE  machineid=@machid)
", Con);
            cmd.Parameters.AddWithValue("@machid", strMacName);
            SqlDataReader rdr;
            Dictionary<string, List<string>> dct = new Dictionary<string, List<string>>();

            try
            {
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string k = rdr["Category"].ToString();
                    string v = rdr["SubCategory"].ToString();
                    if (dct.ContainsKey(k))
                    {
                        dct[k].Add(v);
                    }
                    else
                    {
                        dct.Add(k, new List<string>(new[] { v }));
                    }
                }
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return dct;
        }

        internal static DataTable GetToolLifeData(string strtTime, string endTime, string MachineId)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"s_getToolLifeDetails", Con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@machineid", MachineId);
            cmd.Parameters.AddWithValue("@fromTime", Convert.ToDateTime(strtTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@ToTime", Convert.ToDateTime(endTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@Param", "ScheduledReport");
            SqlDataReader rdr;
            DataTable dt = new DataTable();

            try
            {
                rdr = cmd.ExecuteReader();
                dt.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return dt;
        }

        internal static DataTable GetEWSOEEData(string strtTime, string endTime, string MachineId, string plantid)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            //SqlCommand cmd = new SqlCommand(@"s_getEWSOEE1", Con); // dummy
            SqlCommand cmd = new SqlCommand(@"s_getEWSOEE", Con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 0;
            cmd.Parameters.AddWithValue("@machineid", MachineId);
            cmd.Parameters.AddWithValue("@StartTime", Convert.ToDateTime(strtTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@EndTime", Convert.ToDateTime(endTime).ToString("yyyy-MM-dd HH:mm:ss"));
            SqlDataReader rdr;
            DataTable dt = new DataTable();

            try
            {
                rdr = cmd.ExecuteReader();
                dt.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return dt;
        }

        internal static DataTable GetJHTransactionDetails(string startTime, string endTime, string machine)
        {
            DataTable dtJHTransaction = new DataTable();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            string sProcedure = @"S_View_JHChecklistDashboard_Advik";
            try
            {
                cmd = new SqlCommand(sProcedure, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Machine", machine);
                cmd.Parameters.AddWithValue("@StartDate", startTime);
                cmd.Parameters.AddWithValue("@Enddate", endTime);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    dtJHTransaction.Load(rdr);
                }
            }
            catch(Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (rdr != null) rdr.Close();
                if (conn != null) conn.Close();
            }
            return dtJHTransaction;
        }

        internal static double GetToolLifeThreshold()
        {
            double thresh = 80;
            SqlConnection Con = ConnectionManager.GetConnection();
            //SqlCommand cmd = new SqlCommand(@"s_getEWSOEE1", Con); // dummy
            SqlCommand cmd = new SqlCommand(@"SELECT ValueInInt FROM CockpitDefaults WHERE Parameter='ToolLifeThreshold'", Con);
            SqlDataReader rdr;
            DataTable dt = new DataTable();

            try
            {
                rdr = cmd.ExecuteReader();
                if (rdr.Read())
                {
                    thresh = double.Parse(rdr["ValueInInt"].ToString());
                }
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return thresh;
        }

        internal static DataTable GetEWSWeeklyOEEData(string strtTime, string endTime, string MachineId, string plantid)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetEWSWeeklyOEE]", Con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 0;
            cmd.Parameters.AddWithValue("@machineid", MachineId);
            cmd.Parameters.AddWithValue("@PlantID", plantid);
            cmd.Parameters.AddWithValue("@StartTime", Convert.ToDateTime(strtTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@EndTime", Convert.ToDateTime(endTime).ToString("yyyy-MM-dd HH:mm:ss"));
            SqlDataReader rdr;
            DataTable dt = new DataTable();

            try
            {
                rdr = cmd.ExecuteReader();
                dt.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return dt;
        }

        internal static DataTable GetProductionAndDowntimes(string strtTime, string endTime, string MachineId, string plantid, string parameter, out DataTable dtMachinelist)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetProductionandDowntimeDetails]", Con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@Machineid", MachineId);
            cmd.Parameters.AddWithValue("@plantID", plantid);
            cmd.Parameters.AddWithValue("@StartTime", Convert.ToDateTime(strtTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@EndTime", Convert.ToDateTime(endTime).ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@param", parameter);
            cmd.Parameters.AddWithValue("@machinelist", parameter.Equals("Summary", StringComparison.OrdinalIgnoreCase) ? "Y" : string.Empty);
            SqlDataReader rdr;
            DataTable dt = new DataTable();
            dtMachinelist = new DataTable();
            try
            {
                rdr = cmd.ExecuteReader();
                if (parameter.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                {
                    dtMachinelist.Load(rdr);
                    //rdr.NextResult();
                    //dt.Load(rdr);
                    //rdr.Close();                                       
                }

                dt.Load(rdr);
                rdr.Close();


            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }

            return dt;
        }

        internal static DataTable GetPMReportData(string cell, DateTime startTime, string machine, string plnt, DateTime endTime)
        {
            DataTable dtPMReport = new DataTable();
            SqlConnection conn = null;
            SqlDataReader rdr = null;
            try
            {
                conn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"S_GetPMTransactionReport_Shanti", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", startTime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDate", endTime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Machine", machine);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    dtPMReport.Load(rdr);
                    dtPMReport.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dtPMReport;
        }

        internal static List<DataTable> GetOEEAndLosstime(string MachineId, string strtTime, string endTime)
        {
            DataTable ret1 = new DataTable();
            DataTable ret2 = new DataTable();
            DataTable ret3 = new DataTable();
            DataTable ret4 = new DataTable();
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetAgg_OEEAndLossTimeReport_TAFE]", Con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@StartDate", strtTime);
            cmd.Parameters.AddWithValue("@MachineID", MachineId);
            SqlDataReader rdr;

            try
            {
                rdr = cmd.ExecuteReader();
                ret1.Load(rdr);
                ret2.Load(rdr);
                ret3.Load(rdr);
                ret4.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            return new List<DataTable> { ret1, ret2, ret3, ret4 };
        }

        internal static DataTable GetmangalDowntime(DateTime fromDate, DateTime toDate, out DataTable toprows, out DataTable bottomrows, out DataTable lastcolumn)
        {
            DataTable data = new DataTable();
            toprows = new DataTable();
            bottomrows = new DataTable();
            lastcolumn = new DataTable();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            try
            {
                cmd = new SqlCommand("S_GetForgingProdAnalysisReport_Mangal", conn);
                cmd.Parameters.AddWithValue("@StartTime", fromDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndTime", toDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.CommandType = CommandType.StoredProcedure;
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    toprows.Load(rdr);
                    toprows.AcceptChanges();
                    bottomrows.Load(rdr);
                    bottomrows.AcceptChanges();
                    data.Load(rdr);
                    data.AcceptChanges();
                    lastcolumn.Load(rdr);
                    lastcolumn.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return data;
        }

        internal static void GetEfficiencyAndGraphReport(string machineid, DateTime starttime, DateTime endtime, out DataTable ret1, out DataTable ret2, out DataTable ret3)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            ret1 = new DataTable();
            ret2 = new DataTable();
            ret3 = new DataTable();

            SqlCommand cmd = new SqlCommand(@"[s_GetShiftwiseProdReportFromAutodata_BaluAuto]", Con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@StartDate", starttime.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@EndDate", endtime.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@MachineID", machineid);
            cmd.CommandTimeout = 0;
            SqlDataReader rdr;

            try
            {
                rdr = cmd.ExecuteReader();
                ret1.Load(rdr);
                ret2.Load(rdr);
                ret3.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
        }

        internal static void GetEfficiencyAndGraphReportMonthly(DateTime startTime, out DataTable ret1, out DataTable ret2, out DataTable ret3, out DataTable ret4)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            ret1 = new DataTable();
            ret2 = new DataTable();
            ret3 = new DataTable();
            ret4 = new DataTable();

            SqlCommand cmd = new SqlCommand(@"[s_GetMonthlyProdReportFromAutodata_BaluAuto]", Con);
            cmd.CommandType = CommandType.StoredProcedure;
            //cmd.Parameters.AddWithValue("@StartDate", starttime.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@Date", startTime.ToString("yyyy-MM-dd"));
            //cmd.Parameters.AddWithValue("@Date", "2020-02-14 09:00:00");
            //cmd.Parameters.AddWithValue("@MachineID", machineid);
            cmd.CommandTimeout = 0;
            SqlDataReader rdr;

            try
            {
                rdr = cmd.ExecuteReader();
                ret1.Load(rdr);
                ret2.Load(rdr);
                ret4.Load(rdr);
                ret3.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
        }

        internal static DataTable ShiftProductionCountHourlyBNG(DateTime sttime, string plantname, string machineid)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable dt = new DataTable();
            try
            {
                SqlCommand cmd = new SqlCommand(@"[s_GetHourlyTarget_Count_followup]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Startdate", sttime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDate", sttime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Shift", "");
                cmd.Parameters.AddWithValue("@MachineID", machineid);
                cmd.Parameters.AddWithValue("@param", "BOSCH_BNG_CamShaft");
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    dt.Load(rdr);
                    dt.AcceptChanges();
                }
                dt.Columns.Add("HourLabel");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string hour = gethourminute(dt.Rows[i]["FromTime"].ToString(), dt.Rows[i]["ToTime"].ToString());
                    dt.Rows[i]["HourLabel"] = hour;
                }
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
            return dt;
        }
        internal static string gethourminute(string fromtime, string totime)
        {
            var Fromtime = Convert.ToDateTime(fromtime);
            fromtime = (Fromtime.Minute == 0) ? Fromtime.ToString("HH") : Fromtime.ToString("HH:mm");

            var Totime = Convert.ToDateTime(totime);
            totime = (Totime.Minute == 0) ? Totime.ToString("HH") : Totime.ToString("HH:mm");

            return fromtime + " to " + totime;
        }

        internal static DataTable ShiftProductionCountHourlyBNGAeeLoss(DateTime sttime, string plantname, string machineid)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable dt = new DataTable();
            try
            {
                SqlCommand cmd = new SqlCommand(@"[s_GetHourlyTarget_Count_followup]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@Startdate", sttime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDate", sttime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Shift", "");
                cmd.Parameters.AddWithValue("@MachineID", machineid);
                cmd.Parameters.AddWithValue("@param", "BOSCH_BNG_AELosses");
                SqlDataReader rdr = cmd.ExecuteReader();
                int flag = 0;
                if (rdr.HasRows)
                {
                    dt.Load(rdr);
                    if (dt != null)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(dt.Rows[i]["HourID"]) == 1 && Convert.ToInt32(dt.Rows[i]["ShiftID"]) != 1)
                            {
                                if (flag == 0)
                                {
                                    DataRow toInsert = dt.NewRow();
                                    toInsert[0] = 0;
                                    toInsert[1] = dt.Rows[i]["ShiftID"];
                                    toInsert[2] = dt.Rows[i]["DownCategory"];
                                    toInsert[3] = dt.Rows[i]["DownTime"];
                                    dt.Rows.InsertAt(toInsert, i);
                                    flag = 1;
                                }
                            }
                            else
                            {
                                flag = 0;
                            }
                        }
                        DataRow toInsert1 = dt.NewRow();
                        toInsert1[0] = 0;
                        toInsert1[1] = dt.Rows[dt.Rows.Count - 1]["ShiftID"];
                        toInsert1[2] = dt.Rows[dt.Rows.Count - 1]["DownCategory"];
                        toInsert1[3] = dt.Rows[dt.Rows.Count - 1]["DownTime"];
                        dt.Rows.Add(toInsert1);
                    }
                    dt.AcceptChanges();
                }
                rdr.Close();
            }

            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
            return dt;
        }

        internal static List<string> GetMachineid()
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlDataReader rdr = null;
            List<string> mid = new List<string>();
            try
            {
                SqlCommand cmd = new SqlCommand("select  machineid from machineinformation where TPMTrakEnabled = 1", sqlConn);
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    mid.Add(rdr["machineid"].ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
            return mid;
        }

        internal static void GetSONAMISReportData(DateTime starttime, DateTime endtime, string PlantID, string Shift, out DataTable downTable, out DataTable shiftwiseData)
        {
            //'2019-09-12','2019-09-12'
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            downTable = new DataTable();
            shiftwiseData = new DataTable();
            try
            {
                cmd = new SqlCommand("s_GetSONA_Agg_ShiftwiseProdAndDownReport", sqlConn);
                cmd.CommandTimeout = 180;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", starttime.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@EndDate", endtime.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@MachineID", "");
                cmd.Parameters.AddWithValue("@PlantID", PlantID);
                //cmd.Parameters.AddWithValue("@Shift", Shift);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    downTable.Load(rdr);
                    downTable.AcceptChanges();
                    shiftwiseData.Load(rdr);
                    shiftwiseData.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
                if (rdr != null) rdr.Close();
            }
        }

        internal static List<string> GetAllShift()
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            List<string> shiftList = new List<string>();
            try
            {
                cmd = new SqlCommand(@"select * from shiftDetails where running = 1 order by shiftid", sqlConn);
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    shiftList.Add(rdr["shiftName"].ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
                if (rdr != null) rdr.Close();
            }
            return shiftList;
        }

        internal static DataTable GetFlowMeterReportData(DateTime sttime, DateTime endtime, string plantid, string machineid)
        {
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable dt = new DataTable();
            SqlDataReader rdr = null;
            try
            {
                SqlCommand cmd = new SqlCommand(@"[S_GetBoschFlowMeterReport]", sqlConn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                cmd.Parameters.AddWithValue("@StartDate", sttime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDate", endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@PlantId", plantid);
                cmd.Parameters.AddWithValue("@Mc", machineid);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    dt.Load(rdr);
                    dt.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Error Log - \n " + ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
                if (rdr != null) rdr.Close();
            }
            return dt;
        }

        internal static DataTable GetEnergyReportDataSona(DateTime fromDate, DateTime toDate, string Proc)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            try
            {
                cmd = new SqlCommand(Proc, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@Enddate", toDate.ToString("yyyy-MM-dd"));
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    dt.Load(rdr);
                    dt.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message.ToString());
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dt;
        }

        internal static DataTable GetPlanVsActualData(string plantId, string lineId, string date, out DataTable dtPlanVsActualDataCumulative)
        {
            dtPlanVsActualDataCumulative = new DataTable();
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            DataTable planVsActualDataDaywise = new DataTable();
            SqlDataReader rdr = null;
            try
            {
                SqlCommand cmd = new SqlCommand(@"s_GetTafe_PlanV/sActualReport", sqlConn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", date);

                if (plantId.Equals("All", StringComparison.OrdinalIgnoreCase))
                    cmd.Parameters.AddWithValue("@PlantID", "");
                else
                    cmd.Parameters.AddWithValue("@PlantID", plantId);

                cmd.Parameters.AddWithValue("@Groupid", lineId);
                cmd.CommandTimeout = 120;
                rdr = cmd.ExecuteReader();
                planVsActualDataDaywise.Load(rdr);
                dtPlanVsActualDataCumulative.Load(rdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
                if (rdr != null) rdr.Close();
            }
            return planVsActualDataDaywise;
        }

        internal static DataSet GetOEEAndLosstimeDetails(DateTime fromDate, string machineId)
        {
            DataSet dsOEEAndLosstimeDetails = new DataSet();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlDataAdapter sqlDataAdapter = null;
            try
            {
                SqlCommand sqlCommand = new SqlCommand(@"s_GetTafe_Agg_CategoryWiseOEEAndLossTimeReport", conn);
                sqlCommand.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                sqlCommand.Parameters.AddWithValue("@MachineID", machineId);
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                sqlDataAdapter.Fill(dsOEEAndLosstimeDetails);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (sqlDataAdapter != null) sqlDataAdapter.Dispose();
                if (conn != null) conn.Close();
            }
            return dsOEEAndLosstimeDetails;
        }
        internal static DataTable GetHoldReportData(DateTime fromDate, DateTime toDate, string lineId, string machineId)
        {
            DataTable dtHoldReportData = new DataTable();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlDataReader rdr = null;
            try
            {
                SqlCommand cmd = new SqlCommand("[s_GetTafe_HoldReport]", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@stDateTime", fromDate);
                cmd.Parameters.AddWithValue("@EndDateTime", toDate);
                cmd.Parameters.AddWithValue("@MC", machineId);
                cmd.Parameters.AddWithValue("@GroupId", lineId);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    dtHoldReportData.Load(rdr);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (rdr != null) rdr.Close();
                if (conn != null) conn.Close();
            }
            return dtHoldReportData;
        }

        #region "TAFE Machine History"
        internal static List<MachineHistory> GetMachineHistoryDatas(DateTime fromDateTime, DateTime toDateTime, string machineId)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            List<MachineHistory> machineHistoryData = new List<MachineHistory>();
            SqlDataReader rdr = null;
            string query = @"[s_GetTafe_MachineHistoryViewAndSave]";
            try
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@stDateTime", fromDateTime);
                cmd.Parameters.AddWithValue("@EndDateTime", toDateTime);
                cmd.Parameters.AddWithValue("@MC", machineId);
                cmd.Parameters.AddWithValue("@Param", "View");
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        MachineHistory macHistory = new MachineHistory();
                        macHistory.MachineID = rdr["MachineID"].ToString();
                        macHistory.DownCode = rdr["DownCode"].ToString();
                        macHistory.DownDescription = rdr["Downdescription"].ToString();
                        macHistory.Reason = rdr["Reason"].ToString();
                        macHistory.DownCategory = rdr["DownCategory"].ToString();
                        macHistory.BreakDownStart = Convert.ToDateTime(rdr["BreakDownStart"]).ToString("yyyy-MM-dd hh:mm:ss tt");
                        macHistory.BreakDownEnd = Convert.ToDateTime(rdr["BreakDownEnd"]).ToString("yyyy-MM-dd hh:mm:ss tt");
                        macHistory.ActionProposed = rdr["ActionProposed"].ToString();
                        macHistory.ActionToResolve = rdr["ActionToResolve"].ToString();
                        macHistory.Severity = rdr["Sevierty"].ToString();
                        macHistory.TimeLost = rdr["TimeLost"].ToString();
                        macHistory.ElapsedTime = rdr["ElapsedTime"].ToString();
                        macHistory.KindOfProblem = rdr["KindOfProblem"].ToString();
                        machineHistoryData.Add(macHistory);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (rdr != null) rdr.Close();
                if (conn != null) conn.Close();
            }
            return machineHistoryData;
        }

        internal static bool SaveMachineHistoryData(string machineId, string downCode, string kindOfProb, string downCat, string breakDownStartDate, string reason, string resolveAction, string proposedAction, string severity)
        {
            bool saved = false;
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = @"[s_GetTafe_MachineHistoryViewAndSave]";
            try
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@stDateTime", "");
                cmd.Parameters.AddWithValue("@EndDateTime", "");
                cmd.Parameters.AddWithValue("@MC", machineId);
                cmd.Parameters.AddWithValue("@DownCode", downCode);
                cmd.Parameters.AddWithValue("@Downcatagory", downCat);
                cmd.Parameters.AddWithValue("@BreakDownStartTime", breakDownStartDate);
                cmd.Parameters.AddWithValue("@Reason", reason);
                cmd.Parameters.AddWithValue("@ActionToResolve", resolveAction);
                cmd.Parameters.AddWithValue("@ActionProposed", proposedAction);
                cmd.Parameters.AddWithValue("@Sevierty", severity);
                cmd.Parameters.AddWithValue("@Param", "Save");
                cmd.Parameters.AddWithValue("@KindOfProblem", kindOfProb);
                cmd.CommandType = CommandType.StoredProcedure;
                int cont = cmd.ExecuteNonQuery();
                if (cont > 0)
                {
                    saved = true;
                }
            }
            catch (Exception ex)
            {
                saved = false;
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return saved;
        }

        internal static DataTable GetDowntimeQualificationData(string plnt, string machine, string cell, DateTime fromDate, DateTime toDate)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader reader = null;
            string Query = @"Select A.id,M.machineid,D.downid,A.sttime,A.ndtime,A.mc as MachineInterfaceid,A.dcode as Downinterfaceid from Autodata A
                            inner join machineinformation M on A.mc=M.InterfaceID
                            Left Outer join PlantMachine P on M.machineid=P.MachineID
                            Left outer join PlantMachineGroups PMG on P.PlantID=PMG.PlantID and P.MachineID=PMG.MachineID
                            inner join downcodeinformation D on A.dcode=D.interfaceid
                            where (M.machineid=@machineid or isnull(@machineid,'')='')
                            and (P.PlantID=@Plantid or isnull(@Plantid,'')='') and (PMG.GroupID=@GroupID or isnull(@GroupID,'')='')
                            and (A.sttime>=@sttime and A.ndtime<=@ndtime) and A.datatype=2 and D.downid='NO_DATA'
                            Order by M.machineid,A.sttime";
            try
            {
                cmd = new SqlCommand(Query, conn);
                cmd.Parameters.AddWithValue("@Plantid", plnt);
                cmd.Parameters.AddWithValue("@machineid", machine);
                cmd.Parameters.AddWithValue("@GroupID", cell);
                cmd.Parameters.AddWithValue("@sttime", fromDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@ndtime", toDate.ToString("yyyy-MM-dd HH:mm:ss"));
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    dt.Load(reader);
                }
            }
            catch(Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (reader != null) reader.Close();
                if (conn != null) conn.Close();
            }
            return dt;
        }
        #endregion

        internal static DataTable GetRejectionReportData(DateTime fromDate, DateTime toDate, string plantID, string lineID, string category)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            DataTable dtRejection = new DataTable();
            try
            {
                cmd = new SqlCommand("s_GetTafe_MaterialAndProcessRejectionReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@stDateTime", fromDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDateTime", toDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Plant", plantID);
                cmd.Parameters.AddWithValue("@GroupId", lineID);
                cmd.Parameters.AddWithValue("@Category", category);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                    dtRejection.Load(rdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dtRejection;
        }

        internal static DataTable GetBatchWiseGraphDateReport(DateTime fromDate, string plantID, string lineID, string PartID, string category)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            DataTable dtBatchwiseGraphdata = new DataTable();
            try
            {
                cmd = new SqlCommand("s_GetTafe_BatchwiseReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@PlantID", plantID);
                cmd.Parameters.AddWithValue("@PartID", PartID);
                cmd.Parameters.AddWithValue("@Groupid", lineID);
                cmd.Parameters.AddWithValue("@Catagory", category);
                cmd.Parameters.AddWithValue("@param", "Graph");
                rdr = cmd.ExecuteReader();
                dtBatchwiseGraphdata.Load(rdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dtBatchwiseGraphdata;
        }

        internal static DataTable GetBatchWiseDataReport(DateTime fromDate, string plantID, string lineID, string PartID, string category)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            DataTable dtBatchwisedata = new DataTable();
            try
            {
                cmd = new SqlCommand("s_GetTafe_BatchwiseReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@PlantID", plantID);
                cmd.Parameters.AddWithValue("@PartID", PartID);
                cmd.Parameters.AddWithValue("@Groupid", lineID);
                cmd.Parameters.AddWithValue("@Catagory", category);
                cmd.Parameters.AddWithValue("@param", "");
                rdr = cmd.ExecuteReader();
                dtBatchwisedata.Load(rdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dtBatchwisedata;
        }

        internal static string Getdescription(string partID)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            string partdescription = string.Empty;
            try
            {
                cmd = new SqlCommand("select description from componentinformation where componentid=@compid", conn);
                cmd.Parameters.AddWithValue("@compid", partID);
                rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        partdescription = rdr["description"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return partdescription;
        }

        #region Tafe Line Meter Report
        internal static DataTable GetLinemeterData(string LineID, string starttime, string endtime)
        {

            DataTable ret = new DataTable();
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(@"[s_GetLineMeterGraph_Web_TAFE]", Con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@StartDate", starttime);
            cmd.Parameters.AddWithValue("@EndDate", endtime);
            cmd.Parameters.AddWithValue("@LineID", LineID);
            SqlDataReader rdr = null;

            try
            {
                rdr = cmd.ExecuteReader();
                ret.Load(rdr);
                rdr.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null) Con.Close();
                if (rdr != null) rdr.Close();
            }
            return ret;
        }
        #endregion

        internal static string GellogicalmonthEnd(DateTime fromDate)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalMonth( '" + fromDate.ToString("yyyy-MM-dd 13:00:00") + "','end')", Con);

            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }

        internal static string Gellogicalmonthstart(DateTime fromDate)
        {

            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalMonth( '" + fromDate.ToString("yyyy-MM-dd 13:00:00") + "','start')", Con);

            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return string.Empty;
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }
        internal static DataTable GetComponentDetailsReport(DateTime FromDate, DateTime ToDate, string MachineID)
        {
            DataTable dt = new DataTable();
            SqlConnection connection = ConnectionManager.GetConnection();
            SqlDataReader rdr = null;
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("s_GetL&T_ComponentDetailsReport", connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", FromDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@EndDate", ToDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@MachineID", MachineID = MachineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : MachineID);
                //cmd.Parameters.AddWithValue("@Componentid", ComponentID = ComponentID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : ComponentID);
                //cmd.Parameters.AddWithValue("@OperationNo", OperationNo = OperationNo.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : OperationNo);
                rdr = cmd.ExecuteReader();
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (connection != null) connection.Close();
                if (rdr != null) rdr.Close();
            }
            return dt;
        }
        internal static DataTable GetWeeklyChklistReportData(string machineID, int year)
        {
            SqlConnection Conn = null;
            DataTable dtWeeklyChklistReportData = new DataTable();
            try
            {
                Conn = ConnectionManager.GetConnection();
                SqlCommand command = new SqlCommand(@"s_GetWeekly_TransactionCheckListDetails_GEA", Conn);
                command.Parameters.AddWithValue("@Param", "Report");
                command.Parameters.AddWithValue("@Line", string.Empty);
                command.Parameters.AddWithValue("@Machine", machineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? string.Empty : machineID);
                command.Parameters.AddWithValue("@Date", new DateTime(year, 01, 01));
                command.Parameters.AddWithValue("@freqid", "");
                command.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(dtWeeklyChklistReportData);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (Conn != null) Conn.Close();
            }
            return dtWeeklyChklistReportData;
        }

        internal static DataTable GetDailyChecklistReportData(string lineID, string machineID, string startTime)
        {
            SqlConnection Conn = null;
            DataTable dtDailyChklistReportData = new DataTable();
            try
            {
                Conn = ConnectionManager.GetConnection();
                SqlCommand command = new SqlCommand(@"s_GetDailyCheckListDetails_GEA", Conn);
                command.Parameters.AddWithValue("@Param", "Report");
                command.Parameters.AddWithValue("@Line", lineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? string.Empty : lineID);
                command.Parameters.AddWithValue("@Machine", machineID.Equals("All", StringComparison.OrdinalIgnoreCase) ? string.Empty : machineID);
                command.Parameters.AddWithValue("@Startdate", startTime);
                command.Parameters.AddWithValue("@Frequency", "Daily");
                command.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(dtDailyChklistReportData);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (Conn != null) Conn.Close();
            }
            return dtDailyChklistReportData;
        }

        public static DataTable MachineDownTimeMatrix(DateTime StartDate, DateTime EndDate, string MachineID, string PlantID, string DownID, int Exclude, string MatrixType, string cellId, string proc, string Type)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            DataTable values = new DataTable();
            SqlCommand cmd = new SqlCommand(proc, Con);
            cmd.CommandTimeout = 600;
            if (proc.Equals("s_GetDownTimeMatrixfromAutoData", StringComparison.OrdinalIgnoreCase) || proc.Equals("s_GetSONA_BreakDownMatrix", StringComparison.OrdinalIgnoreCase))
            {
                cmd.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = StartDate.ToString("yyyy-MM-dd HH:mm:ss");
                cmd.Parameters.Add("@EndTime", SqlDbType.NVarChar).Value = EndDate.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else if (proc.Equals("s_GetSONA_ShiftAgg_DowntimeMatrix", StringComparison.OrdinalIgnoreCase) || proc.Equals("s_GetSONA_AggBreakDownMatrix", StringComparison.OrdinalIgnoreCase))
            {
                cmd.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = StartDate.ToString("yyyy-MM-dd");
                cmd.Parameters.Add("@EndTime", SqlDbType.NVarChar).Value = EndDate.ToString("yyyy-MM-dd");
            }
            //cmd.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = StartDate.ToString("yyyy-MM-dd HH:mm:ss");
            //cmd.Parameters.Add("@EndTime", SqlDbType.NVarChar).Value = EndDate.ToString("yyyy-MM-dd HH:mm:ss");
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@PlantID", SqlDbType.NVarChar).Value = PlantID;
            if (proc.Equals("s_GetSONA_BreakDownMatrix", StringComparison.OrdinalIgnoreCase) || proc.Equals("s_GetSONA_AggBreakDownMatrix", StringComparison.OrdinalIgnoreCase))
            {
                if (Type.Equals("ByPhenomenon", StringComparison.OrdinalIgnoreCase))
                    cmd.Parameters.AddWithValue("@BrkDownID", DownID);
                else if (Type.Equals("ByCategory", StringComparison.OrdinalIgnoreCase))
                    cmd.Parameters.AddWithValue("@BrkDownCategory", DownID);
            }
            else
                cmd.Parameters.AddWithValue("@DownID", DownID);
            if (proc.Equals("s_GetSONA_ShiftAgg_DowntimeMatrix", StringComparison.OrdinalIgnoreCase))
                cmd.Parameters.AddWithValue("@Exclude", Exclude);
            else
                cmd.Parameters.AddWithValue("@Excludedown", Exclude);
            cmd.Parameters.AddWithValue("@MatrixType", MatrixType);
            cmd.Parameters.AddWithValue("@Groupid", cellId.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : cellId);
            //cmd.Parameters.AddWithValue("@OperatorID", "");
            //cmd.Parameters.AddWithValue("@ComponentID", "");
            //cmd.Parameters.AddWithValue("@MachineIDLabel", "ALL");
            //cmd.Parameters.AddWithValue("@OperatorIDLabel", "ALL");
            //cmd.Parameters.AddWithValue("@DownIDLabel", "ALL");
            //cmd.Parameters.AddWithValue("@ComponentIDLabel", "ALL");
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader sdr = null;
            try
            {
                sdr = cmd.ExecuteReader();
                values.Load(sdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex);
            }
            finally
            {
                if (Con != null) Con.Close();
            }
            return values;
        }

        public static DataTable AnalysisMachinewiseShiftFormat1Report(DateTime StartDate, string ShiftIn, string MachineID, string ComponentID, string OperationNo, string PlantID, DateTime EndDate, string Param)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            DataTable values = new DataTable();
            SqlCommand cmd = new SqlCommand("s_GetShiftwiseProductionReportFromAutodata", Con);
            cmd.CommandTimeout = 360;
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate.ToString("yyyy-MM-dd");
            cmd.Parameters.Add("@ShiftIn", SqlDbType.NVarChar).Value = ShiftIn;
            cmd.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
            cmd.Parameters.Add("@ComponentID", SqlDbType.NVarChar).Value = ComponentID;
            cmd.Parameters.Add("@OperationNo", SqlDbType.NVarChar).Value = OperationNo;
            cmd.Parameters.Add("@PlantID", SqlDbType.NVarChar).Value = PlantID;
            cmd.Parameters.Add("@EndDate", SqlDbType.NVarChar).Value = EndDate.ToString("yyyy-MM-dd");
            cmd.Parameters.Add("@Param", SqlDbType.NVarChar).Value = Param;
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader sdr = null;
            try
            {
                sdr = cmd.ExecuteReader();
                values.Load(sdr);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception while Executing s_GetShiftwiseProductionReportFromAutodata proc : "+ex.Message);//("GENERATED ERROR : \n" + ex.ToString());
            }
            finally
            {
                if (Con != null) Con.Close();
            }

            return values;
        }

        internal static DataTable GetSAPOEEData(DateTime fromDate, DateTime toDate, string machine, string shift, string plantID, string cell, out DataTable headerDowm)
        {
            DataTable SAPOEEDATA = new DataTable();
            headerDowm = new DataTable();
            SqlDataReader rdr = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            try
            {
                SqlCommand cmd = new SqlCommand(@"s_GetSAPIntegrationReport_Advik", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@EndDate", toDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@ShiftName", shift.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : shift);
                cmd.Parameters.AddWithValue("@PlantID", plantID.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : plantID);
                cmd.Parameters.AddWithValue("@MachineID", machine.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : machine);
                cmd.Parameters.AddWithValue("@CellID", cell.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : cell);
                cmd.Parameters.AddWithValue("@Parameter", "Day");
                rdr = cmd.ExecuteReader();
                headerDowm.Load(rdr);
                headerDowm.AcceptChanges();
                SAPOEEDATA.Load(rdr);
                SAPOEEDATA.AcceptChanges();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return SAPOEEDATA;
        }

        internal static List<string> GetMachinesbyPlantCell(string plant, string cell)
        {
            List<string> MachineList = new List<string>();
            SqlConnection connection = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            string Query = string.Empty;
            if (string.IsNullOrEmpty(plant))
            {
                if (string.IsNullOrEmpty(cell))
                {
                    Query = @"select DISTINCT MachineID from PlantMachineGroups";
                }
                else
                {
                    Query = "select DISTINCT MachineID from PlantMachineGroups where GroupID=@groupID";
                }
            }
            else
            {
                if (string.IsNullOrEmpty(cell))
                {
                    Query = "select DISTINCT MachineID from PlantMachineGroups where PlantID=@plantid";
                }
                else
                {
                    Query = @"select DISTINCT MachineID from PlantMachineGroups where PlantID=@plantid and GroupID=@groupID";
                }
            }
            try
            {
                cmd = new SqlCommand(Query, connection);
                cmd.Parameters.AddWithValue("@plantid", plant);
                cmd.Parameters.AddWithValue("@groupID", cell);
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    MachineList.Add(rdr["MachineID"].ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (connection != null) connection.Close();
                if (rdr != null) rdr.Close();
            }
            return MachineList;
        }

        internal static DataSet Getchecklistdata(string machineID, string shift, DateTime fromDate)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            SqlDataReader rdr = null;
            DataSet dt = new DataSet();
            DataTable shift1val = new DataTable();
            DataTable shift2val = new DataTable();
            DataTable shift3val = new DataTable();
            DataTable shift1Oprsupval = new DataTable();
            DataTable shift2Oprsupval = new DataTable();
            DataTable shift3Oprsupval = new DataTable();
            try
            {
                cmd = new SqlCommand("s_GetJHCheckListReport_Advik", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 300;
                cmd.Parameters.AddWithValue("@MachineID", machineID);
                cmd.Parameters.AddWithValue("@StartDate", fromDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@Shift", shift.Equals("All", StringComparison.OrdinalIgnoreCase) ? "" : shift);
                rdr = cmd.ExecuteReader();
                if (string.IsNullOrEmpty(shift))
                {
                    shift1val.Load(rdr);
                    shift2val.Load(rdr);
                    shift3val.Load(rdr);
                    shift1Oprsupval.Load(rdr);
                    shift2Oprsupval.Load(rdr);
                    shift3Oprsupval.Load(rdr);
                    dt.Tables.Add(shift1val);
                    dt.Tables.Add(shift1Oprsupval);
                    dt.Tables.Add(shift2val);
                    dt.Tables.Add(shift2Oprsupval);
                    dt.Tables.Add(shift3val);
                    dt.Tables.Add(shift3Oprsupval);
                }
                else
                {
                    shift1val.Load(rdr);
                    shift1Oprsupval.Load(rdr);
                    dt.Tables.Add(shift1val);
                    dt.Tables.Add(shift1Oprsupval);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null) conn.Close();
                if (rdr != null) rdr.Close();
            }
            return dt;
        }
    }
}

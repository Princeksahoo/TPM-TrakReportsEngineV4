using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.IO;
using System.Linq;

namespace TPM_TrakReportsEngine
{
    public static class Utility
    {
        public static int WeekNumber(DateTime time)
        {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
   
        public static DateTime FirstDateOfWeek(int year, int weekNum)
        {           
            DateTime jan1 = new DateTime(year, 1, 1);

            int daysOffset = DayOfWeek.Monday - jan1.DayOfWeek;
            DateTime firstMonday = jan1.AddDays(daysOffset);
            Calendar cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstMonday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            if (firstWeek <= 1 || firstWeek > 50)
            {
                weekNum -= 1;
            }
            DateTime result = firstMonday.AddDays(weekNum * 7);
            return result;
        }

        public static void DeleteOldReports(string path, int days)
        {
            try
            {
                List<string> ext = new List<string> { ".pdf", ".xls", ".xlsx" };
                string[] files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))).ToArray<string>();

                foreach (string file in files)
                {
                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        if (fi.LastWriteTime < DateTime.Now.AddDays(-days)) fi.Delete();
                    }
                    catch(Exception ex)
                    {
                        Logger.WriteErrorLog(ex.ToString());
                    }
                }
            }
            catch(Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }
    }
}

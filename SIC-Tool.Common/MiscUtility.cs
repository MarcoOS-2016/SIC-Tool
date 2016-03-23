using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;

namespace SIC_Tool.Common
{
    public class MiscUtility
    {
        private static ILog log = LogManager.GetLogger(typeof(MiscUtility));

        public static void LogHistory(string text)
        {
            string logfilename = "History.log";
            FileUtility.SaveFile(logfilename, string.Format("[{0}] - {1}", DateTime.Now.ToString(), text));
        }

        public static void InsertNewColumn(ref System.Data.DataTable datatablename, string columnname, string columntype, int columnindex)
        {
            DataColumn dc = new DataColumn();
            dc.ColumnName = columnname;
            dc.DataType = System.Type.GetType(columntype);
            datatablename.Columns.Add(dc);

            //DataRow row = null;
            //for (int index = 0; index < datatablename.Tables[0].Rows.Count; index++)
            //{        
            //    row = datatablename.Tables[0].NewRow();
            //    row[columnname] = "";
            //    datatablename.Tables[0].Rows.Add(row);
            //}

            dc.SetOrdinal(columnindex - 1);
        }

        public static void AppendNewColumn(ref List<DataTable> datatablelist, int index, string columnname, string columntype, int columnindex)
        {
            DataColumn dc = new DataColumn();
            dc.ColumnName = columnname;
            dc.DataType = System.Type.GetType(columntype);
            datatablelist[index].Columns.Add(dc);

            dc.SetOrdinal(columnindex - 1);
        }

        public static DataTable CreateNewDataTable(string datatablename, List<string> columnnamelist, List<string> columndatatype)
        {            
            DataTable dt = new DataTable(datatablename);

            DataColumn dc = null;
            for (int index = 0; index < columnnamelist.Count; index++)
            {
                dc = new DataColumn();
                dc.ColumnName = columnnamelist[index];

                if (columndatatype[index].ToUpper().Contains("DATETIME"))
                    dc.DataType = System.Type.GetType("System.DateTime");

                if (columndatatype[index].ToUpper().Contains("STRING"))
                    dc.DataType = System.Type.GetType("System.String");

                if (columndatatype[index].ToUpper().Contains("INT32"))
                    dc.DataType = System.Type.GetType("System.Int32");

                dt.Columns.Add(dc);
            }

            return dt;
        }

        public static string AddSingleQuotation(string sourcestring)
        {
            StringBuilder sb = new StringBuilder();

            foreach (string tempstring in sourcestring.Split(','))
                sb.Append(string.Format("'{0}',", tempstring.Trim()));

            sb.Remove(Convert.ToString(sb).Length - 1, 1);

            return sb.ToString();
        }

        public static string DecryptPassword(string connectionstring)
        {
            string passwordstring = "Password=";

            if (connectionstring.ToUpper().Contains(passwordstring.ToUpper()))
            {
                int startposition = connectionstring.IndexOf(passwordstring) + passwordstring.Length;
                int endposition = connectionstring.LastIndexOf(";");

                string existingpassword = connectionstring.Substring(startposition, endposition - startposition);

                StringBuilder sb = new StringBuilder(connectionstring);
                sb.Replace(existingpassword, PasswordUtility.DesDecrypt(existingpassword));

                return sb.ToString();
            }

            return string.Empty;
        }

        public static string Char2HTML(string sourcestring)
        {
            StringBuilder sb = new StringBuilder(sourcestring);

            sb.Replace("&lt;", "<");
            sb.Replace("&gt;", ">");

            return sb.ToString();
        }

        public static List<string> Sectionalization(List<string> list)
        {
            int limitedlength = 1000;
            List<string> Sectionalizationlist = new List<string>();
            StringBuilder sqlstring = new StringBuilder();

            try
            {
                if (list.Count > 0)
                {
                    if (list.Count <= limitedlength)
                    {
                        for (int count = 0; count < list.Count; count++)
                        {
                            sqlstring.Append("'");
                            sqlstring.Append(list[count]);
                            sqlstring.Append("',");
                        }

                        sqlstring.Remove(sqlstring.Length - 1, 1);
                        Sectionalizationlist.Add(sqlstring.ToString());
                    }
                    else
                    {
                        int multiple = list.Count / limitedlength;
                        int remainder = list.Count % limitedlength;

                        for (int count = 1; count <= multiple * limitedlength; count++)
                        {
                            if (count > 0 && (count % limitedlength == 0))
                            {
                                sqlstring.Remove(sqlstring.Length - 1, 1);
                                Sectionalizationlist.Add(sqlstring.ToString());
                                sqlstring.Clear();
                            }

                            sqlstring.Append("'");
                            sqlstring.Append(list[count - 1]);
                            sqlstring.Append("',");
                        }

                        int startindex = multiple * limitedlength;
                        for (int count = startindex; count < list.Count; count++)
                        {
                            sqlstring.Append("'");
                            sqlstring.Append(list[count]);
                            sqlstring.Append("',");
                        }

                        sqlstring.Remove(sqlstring.Length - 1, 1);
                        Sectionalizationlist.Add(sqlstring.ToString());
                    }
                }

                return Sectionalizationlist;
            }
            catch (System.Exception ex)
            {
                log.Info(ex.Message);
                throw ex;
            }
        }

        public static TimeSpan DateDiff(DateTime datetime1, DateTime datetime2)
        {            
            TimeSpan ts1 = new TimeSpan(datetime1.Ticks);
            TimeSpan ts2 = new TimeSpan(datetime2.Ticks);
            
            return ts1.Subtract(ts2).Duration();
        }
    }    
}

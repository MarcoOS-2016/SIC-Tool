using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;

namespace SIC_Tool.Business
{
    public class UIInqueryReportHandler : ReportHandlerBase
    {
        private string sqlscripttemplate = string.Empty;        
        private SICContext SICContext = null;

        public UIInqueryReportHandler(SICContext SICContext)
        {
            this.SICContext = SICContext;
        }

        public override void Process()
        {
            GetReportSQLScriptTemplate();
            FormatSQLScript();
            FetchRawData();
            SaveReportFile();
        }

        private void GetReportSQLScriptTemplate()
        {
            foreach (UIInqueryReport uiinqueryreport in this.reportconfig.UIInqueryReports)
            {
                if (this.SICContext.UIReport.ReportName.Contains(uiinqueryreport.ReportName.ToUpper()))
                {
                    this.reportname = uiinqueryreport.ReportName;
                    this.databasename = uiinqueryreport.DataBaseName;
                    this.savefolder = this.SICContext.UIReport.SaveFolder;
                    this.sqlscripttemplate = uiinqueryreport.SQLScriptTemplate;
                }
            }
        }

        private void FormatSQLScript()
        {
            StringBuilder sqlstring = new StringBuilder(this.sqlscripttemplate);

            sqlstring.Replace("\r\n", "");

            sqlstring.Replace("%CCN%", FormatSQLString(this.SICContext.UIReport.CCN));
            sqlstring.Replace("%MASLOC%", FormatSQLString(this.SICContext.UIReport.Masloc));
            sqlstring.Replace("%SNAPSHOTDATE%", string.Format("'{0}'", this.SICContext.UIReport.SnapShotDate.ToString("yyyy/MM/dd")));
            sqlstring.Replace("%SNAPSHOTDATE-1%", string.Format("'{0}'", this.SICContext.UIReport.SnapShotDate.AddDays(-1).ToString("yyyy/MM/dd")));

            this.sqlscript = sqlstring.ToString();
        }

        private string FormatSQLString(string characters)
        {
            StringBuilder sb = new StringBuilder();
            if (characters != null && characters.Length != 0)
            {
                foreach (string tempstring in characters.Split(','))
                    sb.Append(string.Format("'{0}',", tempstring.Trim()));

                sb.Remove(Convert.ToString(sb).Length - 1, 1);
            }
            else
            {
                sb.Append("'NULL'");
            }

            return sb.ToString();
        }
    }
}

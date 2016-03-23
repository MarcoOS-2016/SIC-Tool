using System;
using System.IO;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.DataAccess;
using log4net;
using log4net.Config;

namespace SIC_Tool.Business
{
    public abstract class ReportHandlerBase
    {
        protected static readonly ILog log = LogManager.GetLogger(typeof(ReportHandlerBase));
        protected ReportConfig reportconfig = FileUtility.LoadReportConfig();

        protected DataSet rawdata = new DataSet();
        protected string ccn = string.Empty;
        protected string reportname = string.Empty;
        protected string databasename = string.Empty;
        protected string tablename = string.Empty;
        protected string sqlscript = string.Empty;
        protected string savefolder = string.Empty;
        
        public DataSet RawData
        {
            get { return this.rawdata; }
        }

        public ReportHandlerBase()
        {
            XmlConfigurator.Configure(new System.IO.FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log4net.config")));
        }

        public abstract void Process();

        protected virtual void FetchRawData()
        {            
            string connectionstring = string.Empty;

            if (this.databasename.Equals("GLOVIA"))                            
                connectionstring = ConfigFileUtility.GetValue("GloviaDBConnection");

            try
            {
                Console.WriteLine(string.Format("[{0}] - Starting pulling {1} report out from database...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), this.reportname));
                log.Info(string.Format("[{0}] - Starting pulling {1} report out from database...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), this.reportname));

                if (this.databasename.Equals("GLOVIA"))
                {
                    using (OracleAccessDAO dao = new OracleAccessDAO(MiscUtility.DecryptPassword(connectionstring)))
                    {
                        this.rawdata = dao.FetchDataFromDataBase(this.sqlscript);
                    }
                }

                log.Info(string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                //log.Info(string.Format("The Raw data has been pull from {1} database, It will be saved into an excel file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), this.databasename));
                //Console.WriteLine(string.Format("The Raw data has been pull from {1} database, It will be saved into an excel file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), this.databasename));
            }
            catch (Exception e)
            {
                log.Error(string.Format("Function name: [FetchRawData], Error: {0}, StackTrack: {1}", e.Message, e.StackTrace));
                throw;
            }
        }

        protected virtual void SaveReportFile()
        {
            string filename = String.Format("{0}_{1}.csv", this.reportname, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            string fullfilename = Path.Combine(this.savefolder, filename);

            try
            {
                //ExcelFileUtility.SaveExcelFile(fullfilename, rawdata.Tables[0]);
                ExcelFileUtility.ExportDataIntoExcelFile(fullfilename, rawdata.Tables[0]);
                
                log.Info(string.Format("[{0}] - The report file {1} has been created already.", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));
                Console.WriteLine(string.Format("[{0}] - The report file {1} has been created already.", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));
            }
            catch (Exception e)
            {
                log.Error(string.Format("Function name: [SaveReportFile], Error: {0}, StackTrack: {1}", e.Message, e.StackTrace));
                throw;
            }
        }
    }    
}

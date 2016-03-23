using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class SICContext
    {
        private List<string> reportfilenamelist;
        private string vendorreportfilesourcefolder;
        private string vendorreportfileoutputfolder;
        private string reportfilepath;
        private string sourcefolder;
        private string outputfolder;
        private string fbscreportfilepath;
        private List<string> filenamelist;
        private List<string> sheetnamelist;
        private UIReport uireport;
        private DateTime cutoffdate;

        public List<string> ReportFileNameList
        {
            get { return reportfilenamelist; }
            set { reportfilenamelist = value; }
        }

        public string VendorReportFileSourceFolder
        {
            get { return vendorreportfilesourcefolder; }
            set { vendorreportfilesourcefolder = value; }
        }

        public string VendorReportFileOutputFolder
        {
            get { return vendorreportfileoutputfolder; }
            set { vendorreportfileoutputfolder = value; }
        }

        public string ReportFilePath
        {
            get { return reportfilepath; }
            set { reportfilepath = value; }
        }

        public string SourceFolder
        {
            get { return sourcefolder; }
            set { sourcefolder = value; }
        }

        public string OutPutFolder
        {
            get { return outputfolder; }
            set { outputfolder = value; }
        }

        public string FBSCReporFilePath
        {
            get { return fbscreportfilepath; }
            set { fbscreportfilepath = value; }
        }

        public List<string> FileNameList
        {
            get { return filenamelist; }
            set { filenamelist = value; }
        }

        public List<string> SheetNameList
        {
            get { return sheetnamelist; }
            set { sheetnamelist = value; }
        }

        public UIReport UIReport
        {
            get { return uireport; }
            set { uireport = value; }
        }

        public DateTime CutOffDate
        {
            get { return cutoffdate; }
            set { cutoffdate = value; }
        }
    }

    public class UIReport
    {
        private string reportname;
        private DateTime snapshotdate;
        private string ccn;
        private string masloc;
        private string savefolder;

        public string ReportName
        {
            get { return reportname; }
            set { reportname = value; }
        }

        public DateTime SnapShotDate
        {
            get { return snapshotdate; }
            set { snapshotdate = value; }
        }

        public string CCN
        {
            get { return ccn; }
            set { ccn = value; }
        }

        public string Masloc
        {
            get { return masloc; }
            set { masloc = value; }
        }

        public string SaveFolder
        {
            get { return savefolder; }
            set { savefolder = value; }
        }
    }
}

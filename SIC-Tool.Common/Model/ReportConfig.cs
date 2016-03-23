using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class ReportConfig
    {
        private VendorReportFile[] vendorreportfiles;
        private UIInqueryReport[] uiinqueryreports;
        private MappingItem[] mappingitems;

        public VendorReportFile[] VendorReportFiles
        {
            get { return vendorreportfiles; }
            set { vendorreportfiles = value; }
        }

        public UIInqueryReport[] UIInqueryReports
        {
            get { return uiinqueryreports; }
            set { uiinqueryreports = value; }
        }

        public MappingItem[] MappingItems
        {
            get { return mappingitems; }
            set { mappingitems = value; }
        }

        public VendorReportFile this[string reportfilename]
        {
            get
            {
                foreach (VendorReportFile vendorreportfile in vendorreportfiles)
                {
                    if (vendorreportfile.ReportFileName.Contains(reportfilename.Trim().ToUpper()))
                        return vendorreportfile;
                }

                return null;
            }        
        }
    }

    public class VendorReportFile
    {
        private string reportfilename;
        private string keycharsinfilename;
        private string sheetnamelist;
        private string receivingsheetfieldname;
        private string usagesheetfieldname;
        private string nupsheetfieldname;
        private string iasheetfieldname;
        private string reopensheetfieldname;
        private string dcreceivingsheetfieldname;
        private string dcusagesheetfieldname;        

        public string ReportFileName
        {
            get { return reportfilename; }
            set { reportfilename = value; }
        }

        public string KeyCharsInFileName
        {
            get { return keycharsinfilename; }
            set { keycharsinfilename = value; }
        }

        public string SheetNameList
        {
            get { return sheetnamelist; }
            set { sheetnamelist = value; }
        }

        public string ReceivingSheetFieldName
        {
            get { return receivingsheetfieldname; }
            set { receivingsheetfieldname = value; }
        }

        public string UsageSheetFieldName
        {
            get { return usagesheetfieldname; }
            set { usagesheetfieldname = value; }
        }

        public string NUPSheetFieldName
        {
            get { return nupsheetfieldname; }
            set { nupsheetfieldname = value; }
        }

        public string IASheetFieldName
        {
            get { return iasheetfieldname; }
            set { iasheetfieldname = value; }
        }

        public string ReOpenSheetFieldName
        {
            get { return reopensheetfieldname; }
            set { reopensheetfieldname = value; }
        }

        public string DCReceivingSheetFieldName
        {
            get { return dcreceivingsheetfieldname; }
            set { dcreceivingsheetfieldname = value; }
        }

        public string DCUsageSheetFieldName
        {
            get { return dcusagesheetfieldname; }
            set { dcusagesheetfieldname = value; }
        }
    }

    public class UIInqueryReport
    {
        private string reportname;
        private string databasename;
        private string sqlscripttemplate;

        public string ReportName
        {
            get { return reportname; }
            set { reportname = value; }
        }

        public string DataBaseName
        {
            get { return databasename; }
            set { databasename = value; }
        }

        public string SQLScriptTemplate
        {
            get { return sqlscripttemplate; }
            set { sqlscripttemplate = value; }
        }
    }

    public class MappingItem
    {
        private string beginwith;
        private string partner;
        private string kpigoal;
        private string pm;
        private string inventoryanalyst;
        
        public string BeginWith
        {
            get { return beginwith; }
            set { beginwith = value; }
        }

        public string Partner
        {
            get { return partner; }
            set { partner = value; }
        }

        public string KPIGoal
        {
            get { return kpigoal; }
            set { kpigoal = value; }
        }

        public string PM
        {
            get { return pm; }
            set { pm = value; }
        }

        public string InventoryAnalyst
        {
            get { return inventoryanalyst; }
            set { inventoryanalyst = value; }
        }
    }
}

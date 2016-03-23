using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.DataAccess;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;

namespace SIC_Tool.Business.ReportConversion
{
    public abstract class ReportConversionBase
    {
        protected SICContext SICConext = null;
        protected Dictionary<string, PartItem> FBSCdic = new Dictionary<string, PartItem>();

        public ReportConversionBase(SICContext SICContext)
        {
            this.SICConext = SICContext;
        }

        public abstract void Process();

        // Get full sheet name from report file based on defined short sheet name in ReportConfig.xml file.
        protected virtual List<string> GetSheetNames(string fullfilename)
        {
            DataTable sheetnameinfile = null;
            List<string> sheetnamelist = new List<string>();

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fullfilename))
                {
                    sheetnameinfile = dao.GetExcelSheetName();
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));                
                throw;
            }

            string sheetname = string.Empty;
            for (int indey = 0; indey < sheetnameinfile.Rows.Count; indey++)
            {
                if (!sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().Contains("_FilterDatabase"))   //Filter virtual sheet
                    sheetnamelist.Add(sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Replace("$", "").Replace("'", ""));
            }

            return sheetnamelist;
        }

        protected virtual void GetFileNameList()
        {
            this.SICConext.FileNameList = new List<string>();
            DirectoryInfo dir = new DirectoryInfo(SICConext.SourceFolder);

            try
            {
                foreach (FileInfo fi in dir.GetFiles())
                {
                    this.SICConext.FileNameList.Add(fi.FullName);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        protected virtual void ReadFBSCReport()
        {
            DataTable FBSCdatatable = null;

            try
            {
                //string FBSCfilename = ExcelFileUtility.SaveAsStandardFileFormat(SICConext.FBSCReporFilePath, "FBSC");
                string FBSCfilename = SICConext.FBSCReporFilePath;
                PartItem partItem = null;

                List<String> sheetNameList = GetSheetNames(FBSCfilename);

                using (ExcelAccessDAO dao = new ExcelAccessDAO(FBSCfilename))
                {
                    FBSCdatatable = dao.GetPartData(sheetNameList[0]).Tables[0];
                }

                if (FBSCdatatable.Rows.Count != 0)
                {
                    foreach (DataRow dr in FBSCdatatable.Rows)
                    {
                        if (!FBSCdic.ContainsKey(dr["Item"].ToString()))
                        {
                            partItem = new PartItem();
                            partItem.Item = dr["Item"].ToString();
                            partItem.Description = dr["Description"].ToString();
                            partItem.CC = dr["CC"].ToString();
                            partItem.PartCost = Convert.ToDouble(dr["PartCost"]);

                            FBSCdic.Add(dr["Item"].ToString(), partItem);
                        }
                    }
                }

                FBSCdatatable = null;
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        protected virtual void WriteExcelFile(string filename, DataTable[] datatables)
        {
            List<System.Data.DataTable> datatablelist = new List<DataTable>();

            for (int index = 0; index < datatables.Length; index++)
            {
                if (datatables[index] != null)
                    datatablelist.Add(datatables[index]);
            }

            try
            {                
                string fullfilename = Path.Combine(this.SICConext.OutPutFolder,
                    string.Format("{0}_{1}.xlsx", filename, DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                ExcelFileUtility.SaveExcelFileWithMultipleSheets(fullfilename, datatablelist, true);                
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        protected virtual void WriteExcelFile(string filename, List<DataTable> datatablelist)
        {   
            try
            {
                string fullfilename = Path.Combine(this.SICConext.OutPutFolder,
                    string.Format("{0}_{1}.xlsx", filename, DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                ExcelFileUtility.SaveExcelFileWithMultipleSheets(fullfilename, datatablelist, false);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        protected virtual void WriteExcelFile(string filename, DataTable datatable)
        {
            try
            {
                string fullfilename = Path.Combine(this.SICConext.OutPutFolder,
                    string.Format("{0}_{1}.csv", filename, DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                //ExcelFileUtility.SaveExcelFile(fullfilename, datatable, false);
                ExcelFileUtility.ExportDataIntoExcelFile(fullfilename, datatable);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
    }
}

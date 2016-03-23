using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using SIC_Tool.DataAccess;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using log4net;
using log4net.Config;

namespace SIC_Tool.Business
{
    public class InFileHandler
    {
        private static ILog log = LogManager.GetLogger(typeof(InFileHandler));
        public static readonly ReportConfig reportconfig = FileUtility.LoadReportConfig();

        public SICContext SIC_Context = new SICContext();
        public static DataTable ds = new DataTable();
        
        private List<DataTable> receivingdatatablelist = new List<DataTable>();
        private List<DataTable> usagedatatablelist = new List<DataTable>();
        private List<DataTable> nupdatatablelist = new List<DataTable>();
        private List<DataTable> iadatatablelist = new List<DataTable>();
        //private List<DataTable> reopendatatablelist = new List<DataTable>();
        
        public InFileHandler()
        {
            XmlConfigurator.Configure(new System.IO.FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log4net.config")));
        }

        public InFileHandler(SICContext SIC_Context)
        {
            this.SIC_Context = SIC_Context;
        }

        public void Process()
        {
            MiscUtility.LogHistory(string.Format("Searching all report files from the folder - {0}", this.SIC_Context.VendorReportFileSourceFolder));
            CheckReportFiles();
            MiscUtility.LogHistory(string.Format("Total of {0} report files found!", this.SIC_Context.ReportFileNameList.Count));

            string reportfilename = string.Empty;
            DataTable[] datatables = new DataTable[8];

            List<DataSet> datasetlist = new List<DataSet>();
            List<string> fullsheetnamelist = new List<string>();

            try
            {
                for (int index = 0; index < this.SIC_Context.ReportFileNameList.Count; index++)
                {
                    reportfilename = this.SIC_Context.ReportFileNameList[index];

                    for (int indey = 0; indey < reportconfig.VendorReportFiles.Length; indey++)
                    {
                        if (reportfilename.ToUpper().Contains(reportconfig.VendorReportFiles[indey].KeyCharsInFileName.ToUpper()))
                        {
                            MiscUtility.LogHistory(string.Format("Start to get full sheet name list from the excel file - {0}...", reportfilename));
                            fullsheetnamelist = GetFullSheetNames(reportfilename, reportconfig.VendorReportFiles[indey]);
                            MiscUtility.LogHistory("Done!");

                            MiscUtility.LogHistory(string.Format("Start to read the content of excel file - {0}...", reportfilename));
                            ReadReportFileContent(ref datatables, reportfilename, reportconfig.VendorReportFiles[indey], fullsheetnamelist);
                            MiscUtility.LogHistory("Done!");

                            MiscUtility.LogHistory(string.Format("Start to modify the content of excel file - {0}...", reportfilename));
                            ModifyContent(reportfilename, ref datatables);
                            MiscUtility.LogHistory("Done!");
                                                        
                            AppendExcelFile(reportfilename, datatables);
                            
                            //if (reportfilename.ToUpper().Contains("UNI_"))
                            //{
                            //    ds = datatables[0];
                            //}

                            // datasets[0] maps to sheet name - "Receiving *" in 3PL report
                            if (datatables[0] != null)
                                receivingdatatablelist.Add(datatables[0]);
                            
                            // datasets[1] maps to sheet name - "Usage *" in 3PL report
                            if (datatables[1] != null)
                                usagedatatablelist.Add(datatables[1]);

                            // datasets[2] maps to sheet name - "NUP" in 3PL report
                            if (datatables[2] != null)                            
                                nupdatatablelist.Add(datatables[2]);

                            // dataset[3] maps to sheet name - "IA Report" in 3PL report
                            if (datatables[3] != null)
                                iadatatablelist.Add(datatables[3]);

                            for (int indez = 0; indez < datatables.Length; indez++)
                            {
                                datatables[indez] = null;
                            }
                        }
                    }
                }

                MiscUtility.LogHistory("Start to write raw data into the 3PL_Consolidation_Report.xlsx file...");                
                WriteExcelFile();
                MiscUtility.LogHistory("Done!");
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void CheckReportFiles()
        {
            try
            {
                SIC_Context.ReportFileNameList = new List<string>();
                SIC_Context.ReportFileNameList = FileUtility.GetFileNameList(this.SIC_Context.VendorReportFileSourceFolder);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        // Get full sheet name from report file based on defined short sheet name in ReportConfig.xml file.
        private List<string> GetFullSheetNames(string reportfilename, VendorReportFile vendorreportfile)
        {
            DataTable sheetnameinfile = null;
            List<string> fullsheetnamelist = new List<string>();
            string[] shortsheetnames = vendorreportfile.SheetNameList.Split(',');
            
            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                {
                    sheetnameinfile = dao.GetExcelSheetName();
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }

            for (int index = 0; index < shortsheetnames.Length; index++)
            {
                for (int indey = 0; indey < sheetnameinfile.Rows.Count; indey++)
                {
                    //Filter sheet 'Sheetname$filterdatabase', '*_MWD', 'Cancel or Update RMA' from excel file
                    if ((sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains(shortsheetnames[index].ToUpper()))
                        && !sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains("FILTERDATABASE")
                        && !sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains("MWD")
                        && !sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains("CANCEL"))
                        //&& !sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains("DC"))
                    {
                        bool flag = false;
                        string fullsheetname = sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper();

                        //Filter duplicate sheetname
                        for (int indez = 0; indez < fullsheetnamelist.Count; indez++)
                        {
                            if (fullsheetnamelist[indez].Equals(fullsheetname))
                                flag = true;
                        }

                        if (!flag)
                            fullsheetnamelist.Add(fullsheetname);
                    }
                }
            }
            
            return fullsheetnamelist;
        }

        // Read raw data from report files.
        private void ReadReportFileContent(ref DataTable[] datatables, string reportfilename, 
            VendorReportFile vendorreportfile, List<string> fullsheetnamelist)
        {
            List<DataTable> datatablelist = new List<DataTable>();

            try
            {
                for (int index = 0; index < fullsheetnamelist.Count; index++)
                {
                    string fullsheetname = fullsheetnamelist[index].ToUpper();

                    if (vendorreportfile.ReceivingSheetFieldName.Trim().Length != 0
                        && !fullsheetname.Contains("DC")
                        && fullsheetname.Contains("RECEIVING"))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[0] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.ReceivingSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.ReceivingSheetFieldName, datatables[0].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.UsageSheetFieldName.Trim().Length != 0 
                        && !fullsheetname.Contains("DC")
                        && (fullsheetname.Contains("USAGE") || fullsheetname.Contains("SHIPMENT") || fullsheetname.Contains("RNU") || fullsheetname.Contains("RETURNING")))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[1] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.UsageSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.UsageSheetFieldName, datatables[1].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.NUPSheetFieldName.Trim().Length != 0
                        && fullsheetname.Contains("NUP"))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[2] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.NUPSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.NUPSheetFieldName, datatables[2].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.IASheetFieldName.Trim().Length != 0
                        && fullsheetname.Contains("IA"))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[3] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.IASheetFieldName).Tables[0];

                            // Add "'" to format "Account code" column
                            for (int indey = 0; indey < datatables[3].Rows.Count; indey++)
                            {
                                datatables[3].Rows[indey][10] = string.Format("'{0}", datatables[3].Rows[indey][10]);
                            }

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.IASheetFieldName, datatables[3].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.ReOpenSheetFieldName.Trim().Length != 0
                        && fullsheetname.Contains("RE-OPEN"))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[4] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.ReOpenSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.ReOpenSheetFieldName, datatables[4].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.DCReceivingSheetFieldName.Trim().Length != 0
                        && (fullsheetname.Contains("DC RECEIVING")
                        || fullsheetname.Contains("RCV FROM DC")))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[5] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.DCReceivingSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.DCReceivingSheetFieldName, datatables[5].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }

                    if (vendorreportfile.DCUsageSheetFieldName.Trim().Length != 0
                        && (fullsheetname.Contains("DC USAGE")
                        || fullsheetname.Contains("USAGE TO DC")))
                    {
                        using (ExcelAccessDAO dao = new ExcelAccessDAO(reportfilename))
                        {
                            datatables[6] = dao.ReadExcelFile(fullsheetnamelist[index], vendorreportfile.DCUsageSheetFieldName).Tables[0];

                            string text = string.Format("{0} / {1} / {2} / {3}", reportfilename, fullsheetnamelist[index], vendorreportfile.DCUsageSheetFieldName, datatables[6].Rows.Count);
                            FileUtility.SaveFile("Log.txt", text);
                        }

                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void ModifyContent(string reportfilename, ref DataTable[] datatables)
        {
            if (reportfilename.ToUpper().Contains("SRIE_"))
            {
                ModifySIREReportData(ref datatables);
            }

            if (reportfilename.ToUpper().Contains("TAML_"))
            {
                ModifyTAMLReportData(ref datatables);
            }

            if (reportfilename.ToUpper().Contains("SHIPMENT_SER"))
            {
                ModifyDHLReportData(ref datatables);
                MiscUtility.InsertNewColumn(ref datatables[1], "Defective P/N", "System.String", 7);
                MiscUtility.InsertNewColumn(ref datatables[1], "Remark", "System.String", 11);
            }

            if (reportfilename.ToUpper().Contains("REPORT_POW"))
            {
                ModifyCAEReportData(ref datatables);
            }

            if (reportfilename.ToUpper().Contains("UNI_"))
            {
                MiscUtility.InsertNewColumn(ref datatables[0], "Remarks", "System.String", 10);
                MiscUtility.InsertNewColumn(ref datatables[0], "Usage Part", "System.String", 11);
            }

            ChangeFieldName(ref datatables);
        }

        private void AppendExcelFile(string reportfilename, DataTable[] datatables)
        {
            string sheetname = string.Empty;
            string fullfilename = ConfigFileUtility.GetValue("ReOpenFile");

            //If the Re-Open sheet contains raw data in report files, then append the raw data into Re-Open excel file
            if (datatables[4] != null)
            {
                if (datatables[4].Rows.Count != 0)
                {
                    if (reportfilename.ToUpper().Contains("SRIE_"))
                        sheetname = "SRIE cancel from usage";

                    if (reportfilename.ToUpper().Contains("REPORT_TAML"))
                        sheetname = "TAML";

                    if (reportfilename.ToUpper().Contains("REPORT_POW"))
                        sheetname = "CAE";

                    if (reportfilename.ToUpper().Contains("KEAS_POW_DAILY"))
                        sheetname = "KEAS";

                    if (reportfilename.ToUpper().Contains("UNI_PARTS"))
                        sheetname = "UNI";

                    if (reportfilename.ToUpper().Contains("NOP"))
                        sheetname = "NOP";

                    try
                    {
                        MiscUtility.LogHistory(string.Format("Start to append row data into the content of excel file - {0}...", fullfilename));
                        ExcelFileUtility.AppendDateToExcelFile(fullfilename, sheetname, datatables[4]);
                        MiscUtility.LogHistory("Done!");
                    }
                    catch (Exception ex)
                    {
                        MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                        throw;
                    }
                }
            }

            if (datatables[7] != null)
            {
                if (datatables[7].Rows.Count != 0)
                {
                    // Don't add the table into datatable list if the report name is "SRIE" and its sheet name is "NUP"
                    if (reportfilename.ToUpper().Contains("SRIE_"))
                    //&& reportconfig.VendorReportFiles[indey].SheetNameList.Contains("NUP"))
                    {
                        sheetname = "SRIE";

                        try
                        {
                            MiscUtility.LogHistory(string.Format("Start to append row data into the content of excel file - {0}...", fullfilename));
                            ExcelFileUtility.AppendDateToExcelFile(fullfilename, sheetname, datatables[7]);
                            MiscUtility.LogHistory("Done!");
                        }
                        catch (Exception ex)
                        {
                            MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                            throw;
                        }
                    }
                }
            }
        }

        private void WriteExcelFile()
        {
            List<DataTable> datatablelist = MergeReportData();

            try
            {
                string fullfilename = Path.Combine(SIC_Context.VendorReportFileOutputFolder,
                    string.Format("3PL_Consolidation_Report_{0}.xlsx", DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                ExcelFileUtility.SaveExcelFileWithMultipleSheets(fullfilename, datatablelist, false);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
        
        //Uniform all field names in report files as the same field anmes
        private void ChangeFieldName(ref DataTable[] datatables)
        {
            string columnname = string.Empty;
            string datatype = string.Empty;
            List<string[]> fieldnamelist = new List<string[]>();

            try
            {
                foreach (VendorReportFile vendorreportfile in reportconfig.VendorReportFiles)
                {
                    if (vendorreportfile.ReportFileName.ToUpper().Contains("3PL_CONSOLIDATION_REPORT"))
                    {
                        fieldnamelist.Add(vendorreportfile.ReceivingSheetFieldName.Split(','));
                        fieldnamelist.Add(vendorreportfile.UsageSheetFieldName.Split(','));
                        fieldnamelist.Add(vendorreportfile.NUPSheetFieldName.Split(','));
                        fieldnamelist.Add(vendorreportfile.IASheetFieldName.Split(','));
                    }
                }

                for (int indey = 0; indey < fieldnamelist.Count; indey++)
                {
                    if (datatables[indey] != null)
                    {
                        for (int index = 0; index < fieldnamelist[indey].Length; index++)
                        {
                            columnname = fieldnamelist[indey][index].Substring(0, fieldnamelist[indey][index].IndexOf("["));
                            datatables[indey].Columns[index].ColumnName = columnname;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void ModifySIREReportData(ref DataTable[] datatables)
        {
            datatables[7] = datatables[2].Clone();
            int recordcount = datatables[2].Rows.Count;

            int index;
            for (index = 0; index < recordcount; index++)
            {
                // If the column I(column name - "PTA Nr/LPTA Nr") without data in the SRIE report file, then insert into data table.
                if (datatables[2].Rows[index]["PTA Nr/LPTA Nr"].ToString().Trim().Length != 0)
                {
                    datatables[7].Rows.Add(datatables[2].Rows[index].ItemArray);
                    datatables[2].Rows[index].Delete();
                }
            }

            //Remove the last column - "PTA Nr/LPTA Nr"
            datatables[2].Columns.RemoveAt(7);
            datatables[2].AcceptChanges();

            //Combine both sheets - "Replenished Receiving Report" & "DC Receiving Report", datatables[5] contains sheet "DC Receiving Report"            
            for (index = 0; index < datatables[5].Rows.Count; index++)
            {
                datatables[0].Rows.Add(datatables[5].Rows[index].ItemArray);
            }

            //Combine both sheets - "Usage Report" & "DC Usage Report", datatables[6] contains sheet "DC Usage Report"
            for (index = 0; index < datatables[6].Rows.Count; index++)
            {
                datatables[1].Rows.Add(datatables[6].Rows[index].ItemArray);
            }
        }

        private void ModifyTAMLReportData(ref DataTable[] datatables)
        {
            int index;
            //Combine both sheets - "Replenished Receiving Report" & "Rcv from DC", datatables[5] contains sheet "DC Receiving Report"            
            for (index = 0; index < datatables[5].Rows.Count; index++)
            {
                datatables[0].Rows.Add(datatables[5].Rows[index].ItemArray);
            }

            //Combine both sheets - "Usage_Report" & "Usage to DC", datatables[6] contains sheet "DC Usage Report"
            for (index = 0; index < datatables[6].Rows.Count; index++)
            {
                datatables[1].Rows.Add(datatables[6].Rows[index].ItemArray);
            }
        }

        private void ModifyDHLReportData(ref DataTable[] datatables)
        {
            int recordcount = datatables[1].Rows.Count;

            for (int index = 0; index < recordcount; index++)
            {
                if (datatables[1].Rows[index]["Service Type"].ToString().ToUpper().Equals("UPSELL"))
                {
                    datatables[1].Rows[index]["Dsp Name"] = datatables[1].Rows[index]["HuB Code"];                    
                }

                datatables[1].Rows[index][" Used Date "] =
                        DateTime.Parse(datatables[1].Rows[index][" Used Date "].ToString()).ToString("yyyy-MM-dd");
            }
        }

        private void ModifyCAEReportData(ref DataTable[] datatables)
        {
            int recordcount = datatables[1].Rows.Count;

            for (int index = 0; index < recordcount; index++)
            {
                // If the column I(column name - "PTA Nr /LPTA Nr") without data in the Report_POW report file, then delete the row.
                if (datatables[1].Rows[index]["PTA Nr /LPTA Nr"].ToString().Trim().Length == 0)
                {
                    datatables[1].Rows[index].Delete();
                }
            }

            datatables[1].AcceptChanges();
        }

        //Merge all Receiving sheets/Usage sheets/NUP sheets/IA sheets
        private List<DataTable> MergeReportData()
        {
            List<string[]> fieldnamelist = new List<string[]>();

            foreach (VendorReportFile vendorreportfile in reportconfig.VendorReportFiles)
            {
                if (vendorreportfile.ReportFileName.ToUpper().Contains("3PL_CONSOLIDATION_REPORT"))
                {
                    fieldnamelist.Add(vendorreportfile.ReceivingSheetFieldName.Split(','));
                    fieldnamelist.Add(vendorreportfile.UsageSheetFieldName.Split(','));
                    fieldnamelist.Add(vendorreportfile.NUPSheetFieldName.Split(','));
                    fieldnamelist.Add(vendorreportfile.IASheetFieldName.Split(','));
                }
            }            

            int index = 0;
            string columnname = string.Empty;
            string datatype = string.Empty;

            // Create a datatable with table structure of Receiving sheet
            List<string> receivingfieldnamelist = new List<string>();
            List<string> receivingdatatypelist = new List<string>();
            
            for (index = 0; index < fieldnamelist[0].Length; index++)
            {
                columnname = fieldnamelist[0][index].Substring(0, fieldnamelist[0][index].IndexOf("["));
                receivingfieldnamelist.Add(columnname);

                int startposition = fieldnamelist[0][index].IndexOf("[") + 1;
                int length = fieldnamelist[0][index].Length - (startposition + "]".Length);
                datatype = fieldnamelist[0][index].Substring(startposition, length);
                receivingdatatypelist.Add(datatype);
            }

            // Create a datatable with table structure of Usage sheet
            List<string> usagefieldnamelist = new List<string>();
            List<string> usagedatatypelist = new List<string>();

            for (index = 0; index < fieldnamelist[1].Length; index++)
            {
                columnname = fieldnamelist[1][index].Substring(0, fieldnamelist[1][index].IndexOf("["));
                usagefieldnamelist.Add(columnname);

                int startposition = fieldnamelist[1][index].IndexOf("[") + 1;
                int length = fieldnamelist[1][index].Length - (startposition + "]".Length);
                datatype = fieldnamelist[1][index].Substring(startposition, length);
                usagedatatypelist.Add(datatype);
            }

            // Create a datatable with table structure of NUP sheet
            List<string> nupfieldnamelist = new List<string>();
            List<string> nupdatatypelist = new List<string>();

            for (index = 0; index < fieldnamelist[2].Length; index++)
            {
                columnname = fieldnamelist[2][index].Substring(0, fieldnamelist[2][index].IndexOf("["));
                nupfieldnamelist.Add(columnname);

                int startposition = fieldnamelist[2][index].IndexOf("[") + 1;
                int length = fieldnamelist[2][index].Length - (startposition + "]".Length);
                datatype = fieldnamelist[2][index].Substring(startposition, length);
                nupdatatypelist.Add(datatype);
            }

            // Create a datatable with table structure of IA sheet
            List<string> iafieldnamelist = new List<string>();
            List<string> iadatatypelist = new List<string>();

            for (index = 0; index < fieldnamelist[3].Length; index++)
            {
                columnname = fieldnamelist[3][index].Substring(0, fieldnamelist[3][index].IndexOf("["));
                iafieldnamelist.Add(columnname);

                int startposition = fieldnamelist[3][index].IndexOf("[") + 1;
                int length = fieldnamelist[3][index].Length - (startposition + "]".Length);
                datatype = fieldnamelist[3][index].Substring(startposition, length);
                iadatatypelist.Add(datatype);
            }

            try
            {
                DataTable receivingdatatable = MiscUtility.CreateNewDataTable("Receiving Report", receivingfieldnamelist, receivingdatatypelist);
                DataTable usagedatatable = MiscUtility.CreateNewDataTable("Usage Report", usagefieldnamelist, usagedatatypelist);
                DataTable nupdatatable = MiscUtility.CreateNewDataTable("NUP", nupfieldnamelist, nupdatatypelist);
                DataTable iadatatable = MiscUtility.CreateNewDataTable("IA Report", iafieldnamelist, iadatatypelist);

                if (receivingdatatablelist.Count != 0)
                {
                    for (index = 0; index < receivingdatatablelist.Count; index++)
                    {
                        for (int indey = 0; indey < receivingdatatablelist[index].Rows.Count; indey++)
                        {
                            if (receivingdatatablelist[index].Rows[indey][0].ToString().Trim().Length != 0)
                                receivingdatatable.Rows.Add(receivingdatatablelist[index].Rows[indey].ItemArray);
                        }
                    }
                }

                if (usagedatatablelist.Count != 0)
                {
                    for (index = 0; index < usagedatatablelist.Count; index++)
                    {
                        for (int indey = 0; indey < usagedatatablelist[index].Rows.Count; indey++)
                        {
                            if (usagedatatablelist[index].Rows[indey][0].ToString().Trim().Length != 0)
                                usagedatatable.Rows.Add(usagedatatablelist[index].Rows[indey].ItemArray);
                        }
                    }
                }

                if (nupdatatablelist.Count != 0)
                {
                    for (index = 0; index < nupdatatablelist.Count; index++)
                    {
                        for (int indey = 0; indey < nupdatatablelist[index].Rows.Count; indey++)
                        {
                            if (nupdatatablelist[index].Rows[indey][0].ToString().Trim().Length != 0)
                                nupdatatable.Rows.Add(nupdatatablelist[index].Rows[indey].ItemArray);
                        }
                    }
                }

                if (iadatatablelist.Count != 0)
                {
                    for (index = 0; index < iadatatablelist.Count; index++)
                    {
                        for (int indey = 0; indey < iadatatablelist[index].Rows.Count; indey++)
                        {
                            if (iadatatablelist[index].Rows[indey][0].ToString().Trim().Length != 0)
                                iadatatable.Rows.Add(iadatatablelist[index].Rows[indey].ItemArray);
                        }
                    }
                }

                List<DataTable> datatablelist = new List<DataTable>();

                datatablelist.Add(receivingdatatable);
                datatablelist.Add(usagedatatable);
                datatablelist.Add(nupdatatable);
                datatablelist.Add(iadatatable);

                receivingdatatable = null;
                usagedatatable = null;
                nupdatatable = null;
                iadatatable = null;

                return datatablelist;
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
    }
}

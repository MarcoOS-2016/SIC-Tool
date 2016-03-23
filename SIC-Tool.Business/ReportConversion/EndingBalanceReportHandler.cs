using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Text;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.DataAccess;

namespace SIC_Tool.Business.ReportConversion
{
    public class EndingBalanceReportHandler : ReportConversionBase
    {
        DataTable[] datatables = new DataTable[12];
        List<string> filenamelist = new List<string>();        
        List<DataTable> datatablelist = new List<DataTable>();                
        Dictionary<string, List<Int32>> uniquepartnumberdic = new Dictionary<string, List<Int32>>();

        public EndingBalanceReportHandler(SICContext SICContext)
            : base(SICContext)
        {
        }

        ~EndingBalanceReportHandler()
        {
            for (int index = 0; index < datatables.Length; index++)
                datatables[index] = null;

            datatablelist.Clear();
            datatablelist = null;

            uniquepartnumberdic.Clear();
            uniquepartnumberdic = null;
        }

        public override void Process()
        {
            MiscUtility.LogHistory(string.Format("Start to get report file name list from the folder - {0}...", SICConext.SourceFolder));
            GetFileNameList();
            MiscUtility.LogHistory("Done!");

            MiscUtility.LogHistory("Start to read FBSC Report...");
            ReadFBSCReport();
            MiscUtility.LogHistory("Done!");

            for (int index = 0; index < this.SICConext.FileNameList.Count; index++)
            {
                MiscUtility.LogHistory(string.Format("Start to read Ending Balance Report - {0}...", this.SICConext.FileNameList[index]));
                ReadEndingBalanceReport(this.SICConext.FileNameList[index]);
                MiscUtility.LogHistory("Done!");

                MiscUtility.LogHistory(string.Format("Start to modify Ending Balance Report - {0}...", this.SICConext.FileNameList[index]));
                ModifyEndingBalanceReport(this.SICConext.FileNameList[index]);
                MiscUtility.LogHistory("Done!");

                MiscUtility.LogHistory(string.Format("Start to output data into an Excel file - {0}...", this.SICConext.FileNameList[index]));
                string filename = Path.GetFileNameWithoutExtension(this.SICConext.FileNameList[index]);
                WriteExcelFile(filename, datatablelist);
                MiscUtility.LogHistory("Done!");
            }
        }

        private void ReadEndingBalanceReport(string fullfilename)
        {            
            List<string> sheetnamelist = null;
            DataTable datatable = null;
            
            datatablelist.Clear();

            try
            {   
                sheetnamelist = this.GetSheetNames(fullfilename);

                for (int index = 0; index < sheetnamelist.Count; index++)
                {
                    using (ExcelAccessDAO dao = new ExcelAccessDAO(fullfilename))
                    {
                        datatable = new DataTable();
                        datatable = dao.ReadExcelFile(sheetnamelist[index]).Tables[0];
                    }

                    datatable.TableName = sheetnamelist[index];
                    datatablelist.Add(datatable);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        private void ModifyEndingBalanceReport(string fullfilename)
        {   
            RemoveRowsByCondition(fullfilename);
            AddNewColumns(fullfilename);
        }

        // Remove the rows which contains "GICBC" string in the column "BIN" in the sheets which sheetname contains "Ending Balance"
        private void RemoveRowsByCondition(string fullfilename)
        {
            try
            {
                if (fullfilename.ToUpper().Contains("SRIE"))
                {
                    for (int index = 0; index < datatablelist.Count; index++)
                    {
                        if (datatablelist[index].TableName.ToUpper().Contains("ENDING BALANCE"))
                        {   
                            for (int indey = 0; indey < datatablelist[index].Rows.Count; indey++)
                            {
                                if (datatablelist[index].Rows[indey]["BIN"].Equals("GICBC"))
                                {
                                    datatablelist[index].Rows[indey].Delete();
                                }
                            }

                            datatablelist[index].AcceptChanges();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        private void AddNewColumns(string fullfilename)
        {
            //double stardardcost = 0.00;
            PartItem partItem = new PartItem();
            List<Int32> indexnum = null;
            string qtycolumnname = string.Empty;
            
            try
            {
                if (fullfilename.ToUpper().Contains("AIO") || fullfilename.ToUpper().Contains("SBL"))
                    qtycolumnname = "LogicInv";
                else
                    qtycolumnname = "Logic Ending Balance";

                for (int index = 0; index < datatablelist.Count; index++)
                {
                    if (datatablelist[index].TableName.Contains("ENDING BALANCE"))
                    {
                        if (!datatablelist[index].Columns.Contains("STD COST"))
                        {
                            int count = datatablelist[index].Columns.Count + 1;
                            MiscUtility.AppendNewColumn(ref datatablelist, index, "STD COST", "System.Double", count);
                        }

                        if (!datatablelist[index].Columns.Contains("TOTAL COST"))
                        {
                            int count = datatablelist[index].Columns.Count + 1;
                            MiscUtility.AppendNewColumn(ref datatablelist, index, "TOTAL COST", "System.Double", count);
                        }

                        GetUniquePartNumber(index);

                        foreach (KeyValuePair<string, List<Int32>> kvp in uniquepartnumberdic)
                        {
                            if (FBSCdic.TryGetValue(kvp.Key, out partItem))
                            {
                                indexnum = kvp.Value;

                                for (int indey = 0; indey < indexnum.Count; indey++)
                                {
                                    Int16 qty = 0;

                                    if (!string.IsNullOrEmpty(datatablelist[index].Rows[indexnum[indey]][qtycolumnname].ToString()))
                                        qty = Convert.ToInt16(datatablelist[index].Rows[indexnum[indey]][qtycolumnname]);

                                    datatablelist[index].Rows[indexnum[indey]]["STD COST"] = partItem.PartCost;
                                    datatablelist[index].Rows[indexnum[indey]]["TOTAL COST"] = qty * partItem.PartCost;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        private void GetUniquePartNumber(int number)
        {
            string partnumber = string.Empty;
            List<Int32> indexlist = null;

            uniquepartnumberdic.Clear();

            try
            {
                for (int index = 0; index < datatablelist[number].Rows.Count; index++)
                {
                    partnumber = datatablelist[number].Rows[index]["PN"].ToString();

                    if (!uniquepartnumberdic.ContainsKey(partnumber))
                    {
                        indexlist = new List<Int32>();
                        indexlist.Add(index);
                        uniquepartnumberdic.Add(partnumber, indexlist);
                    }
                    else
                    {
                        uniquepartnumberdic[partnumber].Add(index);
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }        
    }
}

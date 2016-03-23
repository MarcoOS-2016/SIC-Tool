using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.DataAccess;

namespace SIC_Tool.Business.ReportConversion
{
    public class LongAgingReportHandler : ReportConversionBase
    {
        public static readonly ReportConfig reportconfig = FileUtility.LoadReportConfig();
                        
        System.Data.DataTable datatable = null;
        Dictionary<string, List<Int32>> uniquepartnumberdic = new Dictionary<string, List<Int32>>();
        
        public LongAgingReportHandler(SICContext SICContext)
            : base(SICContext)
        {
        }
        
        public override void Process()
        {
            ReadReportContent();
            ModifyReportContent();

            MiscUtility.LogHistory("Start to output data into an Excel file...");
            WriteExcelFile("Long Aging Report", datatable);
            MiscUtility.LogHistory("Done!");
        }

        private void ReadReportContent()
        {
            MiscUtility.LogHistory("Start to read FBSC report...");
            ReadFBSCReport();
            MiscUtility.LogHistory("Done!");
            
            MiscUtility.LogHistory("Start to read Long Aging report...");
            ReadLongAgingReport();
            MiscUtility.LogHistory("Done!");
        }

        private void ReadLongAgingReport()
        {
            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(SICConext.ReportFilePath))
                {                    
                    datatable = dao.ReadExcelFile("Sheet1").Tables[0];
                    datatable.TableName = "Sheet1";
                }                
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void ModifyReportContent()
        {
            MiscUtility.LogHistory("Start to add new columns into sheet...");
            AddNewColumns();
            FillCellValue();
            MiscUtility.LogHistory("Done!");
        }

        // Add new columns into the sheet
        private void AddNewColumns()
        {
            try
            {
                MiscUtility.InsertNewColumn(ref datatable, "Description", "System.String", 8);
                MiscUtility.InsertNewColumn(ref datatable, "Commodity", "System.String", 9);
                MiscUtility.InsertNewColumn(ref datatable, "Unit Cost", "System.Double", 10);
                MiscUtility.InsertNewColumn(ref datatable, "Total Cost", "System.Double", 11);
                MiscUtility.InsertNewColumn(ref datatable, "Years", "System.String", 12);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void AssignPartData(ref DataTable datatable, string columnname1, string columnname2, string columnname3)
        {
            PartItem partItem = new PartItem();
            List<Int32> indexnumber = null;

            try
            {
                foreach (KeyValuePair<string, List<Int32>> kvp in uniquepartnumberdic)
                {
                    if (FBSCdic.TryGetValue(kvp.Key, out partItem))
                    {
                        indexnumber = kvp.Value;

                        for (int indey = 0; indey < indexnumber.Count; indey++)
                        {
                            datatable.Rows[indexnumber[indey]][columnname1] = partItem.Description;
                            datatable.Rows[indexnumber[indey]][columnname2] = partItem.CC;
                            datatable.Rows[indexnumber[indey]][columnname3] = partItem.PartCost;
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

        private void FillCellValue()
        {
            GetUniquePartNumber();
            AssignPartData(ref datatable, "Description", "Commodity", "Unit Cost");

            //try
            //{                
                double unitcost = 0.0;
                int qty = 0;
                int days = 0;
                string years = string.Empty;

                for (int index = 0; index < datatable.Rows.Count; index++)
                {
                    //FileUtility.SaveFile("Log.txt",
                    //    string.Format("On Hand Qty - {0}", Convert.ToInt32(datatable.Rows[index]["On Hand Qty"])));

                    if (datatable.Rows[index]["On Hand Qty"].ToString().Length != 0)
                        qty = Convert.ToInt32(datatable.Rows[index]["On Hand Qty"]);
                    else
                        qty = 0;

                    if (datatable.Rows[index]["Unit Cost"].ToString().Length != 0)
                        unitcost = Convert.ToDouble(datatable.Rows[index]["Unit Cost"]);
                    else
                        unitcost = 0.00;

                    // Column name - "Total Cost" in the sheet
                    datatable.Rows[index]["Total Cost"] = qty * unitcost;

                    if (datatable.Rows[index]["diff days"].ToString().Length != 0)
                        days = Convert.ToInt16(datatable.Rows[index]["diff days"]);
                    else
                        days = 0;

                    if (days < 365)
                        years = "<1 year";
                    else if (days >= 365 && days < 2 * 365)
                        years = "1~2 years";
                    else if (days >= 2 * 365 && days < 3 * 365)
                        years = "2~3 years";
                    else if (days >= 3 * 365 && days < 4 * 365)
                        years = "3~4 years";
                    else if (days >= 4 * 365 && days < 5 * 365)
                        years = "4~5 years";
                    else if (days >= 5 * 365 && days < 10 * 365)
                        years = "5~10 years";
                    else if (days >= 10 * 365)
                        years = ">10 years";

                    // Column name - "Years" in the sheet
                    datatable.Rows[index]["Years"] = years;
                }                
            //}
            //catch (Exception ex)
            //{
            //    MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
            //    throw;
            //}
        }
                
        private void GetUniquePartNumber()
        {
            string partnumber = string.Empty;
            List<Int32> indexlist = null;

            uniquepartnumberdic.Clear();

            try
            {
                for (int index = 0; index < datatable.Rows.Count; index++)
                {
                    partnumber = datatable.Rows[index]["Part No"].ToString();

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
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
    }
}

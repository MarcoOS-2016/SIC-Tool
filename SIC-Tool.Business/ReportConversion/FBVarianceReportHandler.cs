using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Text;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.DataAccess;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;

namespace SIC_Tool.Business.ReportConversion
{
    public class FBVarianceReportHandler : ReportConversionBase
    {
        public static readonly ReportConfig reportconfig = FileUtility.LoadReportConfig();

        //System.Data.DataTable stockstatusdatatable = new System.Data.DataTable();
        System.Data.DataTable mappingtable = new System.Data.DataTable();        
        System.Data.DataTable[] datatables = new System.Data.DataTable[5];
        Dictionary<string, List<Int32>> uniquepartnumberdic = new Dictionary<string, List<Int32>>();
        
        public FBVarianceReportHandler(SICContext SICContext)
            : base(SICContext)
        {
        }

        ~FBVarianceReportHandler()
        {
            for (int index = 0; index < datatables.Length; index++)
                datatables[index] = null;

            //stockstatusdatatable.Clear();
            //stockstatusdatatable = null;

            mappingtable.Clear();
            mappingtable = null;
        }

        public override void Process()
        {
            ReadReportContent();
            ModifyReportContent();

            MiscUtility.LogHistory("Start to output data into an Excel file...");
            WriteExcelFile("FB Report", datatables);
            MiscUtility.LogHistory("Done!");
        }

        private void ReadReportContent()
        {
            MiscUtility.LogHistory("Start to read FBSC report...");
            ReadFBSCReport();
            MiscUtility.LogHistory("Done!");

            MiscUtility.LogHistory("Start to read On Way Checking report...");
            ReadFBOriginalReport();
            MiscUtility.LogHistory("Done!");
        }

        private void ReadFBOriginalReport()
        {
            List<string> columnnamelist = new List<string>();
            columnnamelist.Add("Order Nr");
            columnnamelist.Add("SER#");
            ExcelFileUtility.ChangeColumnDataType(SICConext.ReportFilePath, "TO_FB", columnnamelist, "'");

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(SICConext.ReportFilePath))
                {                    
                    datatables[0] = dao.ReadExcelFile("TO_FB").Tables[0];
                    datatables[0].TableName = "FB Inventory";

                    datatables[1] = dao.ReadExcelFile("FB Variance").Tables[0];
                    datatables[1].TableName = "FB Variance GL&Recon";

                    datatables[2] = new DataTable();
                    datatables[2].TableName = "FB Variance Report";

                    datatables[3] = new DataTable();
                    datatables[3].TableName = "3PL Performance";

                    datatables[4] = datatables[0].Copy();
                    datatables[4].TableName = "Raw Data for FB Inventory";
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
            AddNewColumns();

            MiscUtility.LogHistory("Start to build FB Inventory sheet...");
            BuildFBInventorySheet();
            MiscUtility.LogHistory("Done!");

            MiscUtility.LogHistory("Start to build FB Variance GL&Recon sheet...");
            BuildFBVarianceGLReconSheet();
            MiscUtility.LogHistory("Done!");

            MiscUtility.LogHistory("Start to build FB Variance Report sheet...");
            BuildFBVarianceReportSheet();
            MiscUtility.LogHistory("Done!");

            MiscUtility.LogHistory("Start to build 3PL Performance sheet...");
            Build3PLPerformanceSheet();
            MiscUtility.LogHistory("Done!");
        }

        private void BuildFBInventorySheet()
        {
            DeleteConditionalRows();
            FillCellValue();
        }

        // Build 3PL performance sheet
        private void Build3PLPerformanceSheet()
        {
            bool flag = false;
            double totalcostbyitem = 0.00;

            List<AgingItem> agingitemlist = new List<AgingItem>();
            AgingItem tempagingitem = new AgingItem();

            try
            {
                for (int index = 0; index < datatables[0].Rows.Count; index++)
                {
                    flag = false;

                    for (int indey = 0; indey < agingitemlist.Count; indey++)
                    {
                        if (string.Compare(datatables[0].Rows[index]["DSP"].ToString(), agingitemlist[indey].SupplierName, 
                            StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            totalcostbyitem = Convert.ToDouble(datatables[0].Rows[index]["Total Cost"]);

                            switch (datatables[0].Rows[index]["Aging"].ToString())
                            {
                                case "0~7 days":
                                    agingitemlist[indey].Days1 += totalcostbyitem;                                    
                                    break;

                                case "8~20 days":
                                    agingitemlist[indey].Days2 += totalcostbyitem;
                                    break;

                                case "21~30 days":
                                    agingitemlist[indey].Days3 += totalcostbyitem;                                    
                                    break;

                                case "31~60 days":
                                    agingitemlist[indey].Days4 += totalcostbyitem;                                    
                                    break;

                                case ">60 days":
                                    agingitemlist[indey].Days5 += totalcostbyitem;                                    
                                    break;
                            }

                            agingitemlist[indey].AgingTotalCost += totalcostbyitem;

                            flag = true;
                            break;
                        }
                    }

                    if (!flag)
                    {
                        tempagingitem = new AgingItem();
                        tempagingitem.SupplierName = datatables[0].Rows[index]["DSP"].ToString().Trim();
                        totalcostbyitem = Convert.ToDouble(datatables[0].Rows[index]["Total Cost"]);

                        switch (datatables[0].Rows[index]["Aging"].ToString())
                        {
                            case "0~7 days":
                                tempagingitem.Days1 = totalcostbyitem;
                                break;

                            case "8~20 days":
                                tempagingitem.Days2 = totalcostbyitem;
                                break;

                            case "21~30 days":
                                tempagingitem.Days3 = totalcostbyitem;
                                break;

                            case "31~60 days":
                                tempagingitem.Days4 = totalcostbyitem;
                                break;

                            case ">60 days":
                                tempagingitem.Days5 = totalcostbyitem;
                                break;
                        }

                        tempagingitem.AgingTotalCost = totalcostbyitem;
                        agingitemlist.Add(tempagingitem);
                    }                    
                }

                // Add percentage data by Supplier & aging days
                DataRow datarow = null;
                string kpigoal = string.Empty;

                for (int index = 0; index < agingitemlist.Count; index++)
                {
                    for (int indey = 0; indey < reportconfig.MappingItems.Length; indey++)
                    {
                        if (string.Compare(agingitemlist[index].SupplierName.ToString(), reportconfig.MappingItems[indey].Partner, 
                            StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            kpigoal = reportconfig.MappingItems[indey].KPIGoal;
                            break;
                        }
                    }

                    datarow = datatables[3].NewRow();

                    datarow["Supplier"] = agingitemlist[index].SupplierName;
                    datarow["KPI Goal"] = kpigoal;
                    datarow["0~7 Days"] = (agingitemlist[index].Days1 / agingitemlist[index].AgingTotalCost).ToString("p");
                    datarow["7~20 Days"] = (agingitemlist[index].Days2 / agingitemlist[index].AgingTotalCost).ToString("p");
                    datarow["21~30 Days"] = (agingitemlist[index].Days3 / agingitemlist[index].AgingTotalCost).ToString("p");
                    datarow["31~60 Days"] = (agingitemlist[index].Days4 / agingitemlist[index].AgingTotalCost).ToString("p");
                    datarow[">60 Days"] = (agingitemlist[index].Days5 / agingitemlist[index].AgingTotalCost).ToString("p");
                    datarow["On-way Value"] = (agingitemlist[index].Days1 
                        + agingitemlist[index].Days2 
                        + agingitemlist[index].Days3
                        + agingitemlist[index].Days4 
                        + agingitemlist[index].Days5).ToString("C");

                    datatables[3].Rows.Add(datarow);
                }

                // Add Total row
                double totaldays1 = 0.00;
                double totaldays2 = 0.00;
                double totaldays3 = 0.00;
                double totaldays4 = 0.00;
                double totaldays5 = 0.00;

                for (int index = 0; index < agingitemlist.Count; index++)
                {
                    //totaldays1 += agingitemlist[index].Days1 + agingitemlist[index].Days2 + agingitemlist[index].Days3 + agingitemlist[index].Days4 + agingitemlist[index].Days5;
                    totaldays1 += agingitemlist[index].Days1;
                    totaldays2 += agingitemlist[index].Days2;
                    totaldays3 += agingitemlist[index].Days3;
                    totaldays4 += agingitemlist[index].Days4;
                    totaldays5 += agingitemlist[index].Days5;
                }

                double totalagingcost = totaldays1 + totaldays2 + totaldays3 + totaldays4 + totaldays5;
                
                datarow = datatables[3].NewRow();

                datarow["Supplier"] = "TOTAL";
                datarow["KPI Goal"] = "98%";
                datarow["0~7 Days"] = (totaldays1 / totalagingcost).ToString("p");
                datarow["7~20 Days"] = (totaldays2/ totalagingcost).ToString("p");
                datarow["21~30 Days"] = (totaldays3 / totalagingcost).ToString("p");
                datarow["31~60 Days"] = (totaldays4 / totalagingcost).ToString("p");
                datarow[">60 Days"] = (totaldays5 / totalagingcost).ToString("p");
                datarow["On-way Value"] = totalagingcost.ToString("C");

                datatables[3].Rows.Add(datarow);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void BuildFBVarianceReportSheet()
        {
            bool flag = false;            
            List<SupplierItem> supplieritemlist = new List<SupplierItem>();
            SupplierItem tempsupplieritem = null;

            try
            {
                for (int index = 0; index < datatables[1].Rows.Count; index++)
                {
                    flag = false;

                    for (int indey = 0; indey < supplieritemlist.Count; indey++)
                    {
                        if (string.Compare(datatables[1].Rows[index]["Supplier"].ToString(), supplieritemlist[indey].SupplierName, 
                            StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            Int32 onwayinventory = Convert.ToInt32(datatables[1].Rows[index]["ON Way Inventory"]);
                            supplieritemlist[indey].OnWayQty += onwayinventory;

                            Int32 gloviaqty = Convert.ToInt32(datatables[1].Rows[index]["Glovia Qty"]);
                            supplieritemlist[indey].GloviaQty += gloviaqty;

                            Int32 betweenqty = Convert.ToInt32(datatables[1].Rows[index]["Glovia-Access"]);
                            supplieritemlist[indey].BetweenQty += betweenqty;

                            double stdcost = 0.00;
                            if (datatables[1].Rows[index]["STDCOST"].ToString().Length != 0)
                                stdcost = Convert.ToDouble(datatables[1].Rows[index]["STDCOST"]);
                            else
                                stdcost = 0.00;

                            supplieritemlist[indey].OnWayValue += (stdcost * onwayinventory);
                            supplieritemlist[indey].GloviaValue += (stdcost * gloviaqty);
                            supplieritemlist[indey].BetweenValue += (stdcost * betweenqty);                            

                            flag = true;
                            break;
                        }
                    }

                    if (!flag)
                    {
                        tempsupplieritem = new SupplierItem();

                        tempsupplieritem.SupplierName = datatables[1].Rows[index]["Supplier"].ToString();
                        Int32 onwayinventory = Convert.ToInt32(datatables[1].Rows[index]["ON Way Inventory"]);
                        tempsupplieritem.OnWayQty = onwayinventory;

                        Int32 gloviaqty = Convert.ToInt32(datatables[1].Rows[index]["Glovia Qty"]);
                        tempsupplieritem.GloviaQty = gloviaqty;

                        Int32 betweenqty = Convert.ToInt32(datatables[1].Rows[index]["Glovia-Access"]);
                        tempsupplieritem.BetweenQty = betweenqty;

                        double stdcost = 0.00;
                        if (datatables[1].Rows[index]["STDCOST"] != null)
                            stdcost = Convert.ToDouble(datatables[1].Rows[index]["STDCOST"]);
                        else
                            stdcost = 0.00;

                        tempsupplieritem.OnWayValue = (stdcost * onwayinventory);
                        tempsupplieritem.GloviaValue = (stdcost * gloviaqty);
                        tempsupplieritem.BetweenValue = (stdcost * betweenqty);

                        supplieritemlist.Add(tempsupplieritem);                        
                    }
                }

                // Calculate total of items
                Int32 totalonwayqty = 0;
                Int32 totalgloviaqty = 0;
                Int32 totalbetweenqty = 0;
                double totalonwayvalue = 0.00;
                double totalgloviavalue = 0.00;
                double totalbetweenvalue = 0.00;
                //double percentage = 0.00;

                for (int index = 0; index < supplieritemlist.Count; index++)
                {
                    totalonwayqty += supplieritemlist[index].OnWayQty;
                    totalgloviaqty += supplieritemlist[index].GloviaQty;
                    totalbetweenqty += supplieritemlist[index].BetweenQty;

                    totalonwayvalue += supplieritemlist[index].OnWayValue;
                    totalgloviavalue += supplieritemlist[index].GloviaValue;
                    totalbetweenvalue += supplieritemlist[index].BetweenValue;

                    supplieritemlist[index].Percentage = supplieritemlist[index].BetweenValue / supplieritemlist[index].GloviaValue;

                    for (int indey = 0; indey < reportconfig.MappingItems.Length; indey++)
                    {
                        if (string.Compare(supplieritemlist[index].SupplierName, reportconfig.MappingItems[indey].Partner,
                            StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            supplieritemlist[index].Owner = reportconfig.MappingItems[indey].InventoryAnalyst;
                            break;
                        }
                    }
                }
                
                tempsupplieritem = new SupplierItem();

                tempsupplieritem.SupplierName = "Total";
                tempsupplieritem.OnWayQty = totalonwayqty;
                tempsupplieritem.GloviaQty = totalgloviaqty;
                tempsupplieritem.BetweenQty = totalbetweenqty;
                tempsupplieritem.OnWayValue = totalonwayvalue;
                tempsupplieritem.GloviaValue = totalgloviavalue;
                tempsupplieritem.BetweenValue = totalbetweenvalue;
                tempsupplieritem.Percentage = totalbetweenvalue / totalgloviavalue;

                supplieritemlist.Add(tempsupplieritem);

                DataRow datarow = null;
                for (int index = 0; index < supplieritemlist.Count; index++)
                {
                    datarow = datatables[2].NewRow();

                    datarow["Supplier"] = supplieritemlist[index].SupplierName;
                    datarow["On-Way Qty"] = supplieritemlist[index].OnWayQty;
                    datarow["Glovia Qty"] = supplieritemlist[index].GloviaQty;
                    datarow["Gross QTY between GL&Recon"] = Math.Abs(supplieritemlist[index].BetweenQty);
                    datarow["On-Way Value"] = supplieritemlist[index].OnWayValue.ToString("C");
                    datarow["Glovia Value"] = supplieritemlist[index].GloviaValue.ToString("C");
                    datarow["Gross Value between GL&Recon"] = Math.Abs(supplieritemlist[index].BetweenValue).ToString("C");
                    datarow["Percentage"] = Math.Abs(supplieritemlist[index].Percentage).ToString("p");
                    datarow["Owner"] = supplieritemlist[index].Owner;

                    datatables[2].Rows.Add(datarow);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        // Remove rows which delete column C DSPName begin with "MSC" or column E Ser# begin with "1" from raw data
        private void DeleteConditionalRows()
        {
            string text = string.Empty;

            for (int index = 0; index < datatables[0].Rows.Count; index++)
            {
                text = datatables[0].Rows[index]["Dsp name"].ToString().Trim().ToUpper();
                //FileUtility.SaveFile("DataTable.txt", text);
                if (text.Length >= 3)
                {
                    if (text.Substring(0, 3).Equals("MSC"))
                    {
                        datatables[0].Rows[index].Delete();
                        continue;
                    }
                }

                text = datatables[0].Rows[index]["SER#"].ToString().Trim();
                //FileUtility.SaveFile("DataTable.txt", text);
                if (text.Length != 0)
                {
                    if (text.Substring(0, 1).Equals("1"))
                    {
                        datatables[0].Rows[index].Delete();
                    }
                }
            }

            datatables[0].AcceptChanges();
        }

        // Add new columns into FB Inventory sheet
        private void AddNewColumns()
        {
            try
            {
                MiscUtility.InsertNewColumn(ref datatables[0], "DSP", "System.String", 4);
                MiscUtility.InsertNewColumn(ref datatables[0], "Planner", "System.String", 5);
                MiscUtility.InsertNewColumn(ref datatables[0], "Forward Feedback", "System.String", 8);
                MiscUtility.InsertNewColumn(ref datatables[0], "Consignee Feedback", "System.String", 9);
                MiscUtility.InsertNewColumn(ref datatables[0], "Shipper Feedback", "System.String", 10);
                MiscUtility.InsertNewColumn(ref datatables[0], "Inventory Remark", "System.String", 11);
                MiscUtility.InsertNewColumn(ref datatables[0], "Cutoff Date", "System.String", 19);
                MiscUtility.InsertNewColumn(ref datatables[0], "TAT", "System.Int32", 20);
                MiscUtility.InsertNewColumn(ref datatables[0], "Aging", "System.String", 21);
                MiscUtility.InsertNewColumn(ref datatables[0], "Unit Cost", "System.Double", 22);
                MiscUtility.InsertNewColumn(ref datatables[0], "Total Cost", "System.Double", 23);
                MiscUtility.InsertNewColumn(ref datatables[0], "Supplier", "System.String", 24);

                MiscUtility.InsertNewColumn(ref datatables[2], "Supplier", "System.String", 1);
                MiscUtility.InsertNewColumn(ref datatables[2], "On-Way Qty", "System.Int32", 2);
                MiscUtility.InsertNewColumn(ref datatables[2], "Glovia Qty", "System.Int32", 3);
                MiscUtility.InsertNewColumn(ref datatables[2], "Gross QTY between GL&Recon", "System.Int32", 4);
                MiscUtility.InsertNewColumn(ref datatables[2], "On-Way Value", "System.String", 5);
                MiscUtility.InsertNewColumn(ref datatables[2], "Glovia Value", "System.String", 6);
                MiscUtility.InsertNewColumn(ref datatables[2], "Gross Value between GL&Recon", "System.String", 7);
                MiscUtility.InsertNewColumn(ref datatables[2], "Percentage", "System.String", 8);
                MiscUtility.InsertNewColumn(ref datatables[2], "Owner", "System.String", 9);

                MiscUtility.InsertNewColumn(ref datatables[3], "Supplier", "System.String", 1);
                MiscUtility.InsertNewColumn(ref datatables[3], "KPI Goal", "System.String", 2);
                MiscUtility.InsertNewColumn(ref datatables[3], "0~7 Days", "System.String", 3);
                MiscUtility.InsertNewColumn(ref datatables[3], "7~20 Days", "System.String", 4);
                MiscUtility.InsertNewColumn(ref datatables[3], "21~30 Days", "System.String", 5);
                MiscUtility.InsertNewColumn(ref datatables[3], "31~60 Days", "System.String", 6);
                MiscUtility.InsertNewColumn(ref datatables[3], ">60 Days", "System.String", 7);
                MiscUtility.InsertNewColumn(ref datatables[3], "Forwarder", "System.String", 8);
                MiscUtility.InsertNewColumn(ref datatables[3], "On-way Value", "System.String", 9);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void AssignStandardCost(ref DataTable datatable, string columnname)
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
                            datatable.Rows[indexnumber[indey]][columnname] = partItem.PartCost;
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
            TimeSpan timediff;
            DateTime shippingdate;
            string aging = string.Empty;
            string dspnamebeginchar = string.Empty;
            string currentdate = SICConext.CutOffDate.ToString("MM/dd/yyyy");

            GetUniquePartNumber("FB Inventory");
            AssignStandardCost(ref datatables[0], "Unit Cost");

            try
            {
                for (int index = 0; index < datatables[0].Rows.Count; index++)
                {
                    if (datatables[0].Rows[index]["Dsp name"].ToString().Length != 0)
                    {
                        for (int indey = 0; indey < reportconfig.MappingItems.Length; indey++)
                        {
                            dspnamebeginchar = datatables[0].Rows[index]["Dsp name"].ToString().ToUpper().Substring(0, 1);

                            if (string.Compare(dspnamebeginchar, reportconfig.MappingItems[indey].BeginWith, 
                                StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                // Column name - "DSP" in the sheet - "FB Inventory"
                                datatables[0].Rows[index]["DSP"] = reportconfig.MappingItems[indey].Partner.Trim();

                                // Column name - "Planner" in the sheet - "FB Inventory"
                                datatables[0].Rows[index]["Planner"] = reportconfig.MappingItems[indey].PM.Trim();

                                break;
                            }
                        }
                    }                    

                    shippingdate = Convert.ToDateTime(datatables[0].Rows[index]["Shipping date"]);
                    timediff = MiscUtility.DateDiff(SICConext.CutOffDate, shippingdate);

                    // Column name - "Cutoff Date" in the sheet - "FB Inventory"
                    datatables[0].Rows[index]["Cutoff Date"] = currentdate;

                    // Column name - "TAT" in the sheet - "FB Inventory"
                    datatables[0].Rows[index]["TAT"] = timediff.Days.ToString();

                    int daysdiff = Convert.ToInt32(timediff.Days);
                    if (daysdiff >= 0 && daysdiff <= 7)
                        aging = "0~7 days";
                    else if (daysdiff >= 8 && daysdiff <= 20)
                        aging = "8~20 days";
                    else if (daysdiff >= 21 && daysdiff <= 30)
                        aging = "21~30 days";
                    else if (daysdiff >= 31 && daysdiff <= 60)
                        aging = "31~60 days";
                    else if (daysdiff >= 61)
                        aging = ">60 days";

                    // Column name - "Aging" in the sheet - "FB Inventory"
                    datatables[0].Rows[index]["Aging"] = aging;

                    double unitcost = 0.0;
                    int qty = Convert.ToInt32(datatables[0].Rows[index]["FB Inventory"]);

                    if (datatables[0].Rows[index]["Unit Cost"].ToString().Length != 0)
                        unitcost = Convert.ToDouble(datatables[0].Rows[index]["Unit Cost"]);

                    // Column name - "Total Cost" in the sheet - "FB Inventory"
                    datatables[0].Rows[index]["Total Cost"] = qty * unitcost;

                    // Column name - "Supplier" in the sheet - "FB Inventory"
                    string hubcodebeginchar = string.Empty;
                    for (int indey = 0; indey < reportconfig.MappingItems.Length; indey++)
                    {
                        if (datatables[0].Rows[index]["Hub Code"].ToString().Length != 0)
                        {
                            hubcodebeginchar = datatables[0].Rows[index]["Hub Code"].ToString().ToUpper().Substring(0, 1);

                            if (string.Compare(hubcodebeginchar, reportconfig.MappingItems[indey].BeginWith,
                                StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                datatables[0].Rows[index]["Supplier"] = reportconfig.MappingItems[indey].Partner;
                                break;
                            }
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

        // Add new columns - "Supplier, Total Cost" into "FB Variance GL & Recon" sheet
        private void BuildFBVarianceGLReconSheet()
        {
            MiscUtility.InsertNewColumn(ref datatables[1], "Glovia Qty", "System.Int32", 11);
            MiscUtility.InsertNewColumn(ref datatables[1], "Total Cost", "System.Double", 12);
            MiscUtility.InsertNewColumn(ref datatables[1], "Supplier", "System.String", 13);

            GetUniquePartNumber("FB Variance GL&Recon");
            AssignStandardCost(ref datatables[1], "STDCOST");

            try
            {
                for (int index = 0; index < datatables[1].Rows.Count; index++)
                {
                    //int indey = 0;
                    //for (indey = 0; indey < stockstatusdatatable.Rows.Count; indey++)
                    //{
                    //    if (string.Compare(datatables[1].Rows[index]["P/N"].ToString(), stockstatusdatatable.Rows[indey]["ITEM"].ToString(),
                    //        StringComparison.OrdinalIgnoreCase) == 0)
                    //    {
                    //        datatables[1].Rows[index]["STDCOST"] = stockstatusdatatable.Rows[indey]["STD_COST"];
                    //        break;
                    //    }
                    //}

                    string hubcodebeginchar = string.Empty;
                    for (int indey = 0; indey < reportconfig.MappingItems.Length; indey++)
                    {
                        if (datatables[1].Rows[index]["HubCode"].ToString().Length != 0)
                        {
                            hubcodebeginchar = datatables[1].Rows[index]["HubCode"].ToString().ToUpper().Substring(0, 1);

                            if (string.Compare(hubcodebeginchar, reportconfig.MappingItems[indey].BeginWith,
                                StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                datatables[1].Rows[index]["Supplier"] = reportconfig.MappingItems[indey].Partner;
                                break;
                            }
                        }
                    }

                    double stdcost = 0.0;
                    int qty = Convert.ToInt32(datatables[1].Rows[index]["ON Way Inventory"]);

                    if (datatables[1].Rows[index]["STDCOST"].ToString().Length != 0)
                        stdcost = Convert.ToDouble(datatables[1].Rows[index]["STDCOST"]);

                    datatables[1].Rows[index]["Total Cost"] = qty * stdcost;
                    datatables[1].Rows[index]["Glovia Qty"] =
                        Convert.ToInt32(datatables[1].Rows[index]["FB OH_QTY"]) 
                        + Convert.ToInt32(datatables[1].Rows[index]["PTA/LPTA To FB Pending"])
                        - Convert.ToInt32(datatables[1].Rows[index]["Receiving Pending"])
                        - Convert.ToInt32(datatables[1].Rows[index]["XLOC Revise From Pending"])
                        + Convert.ToInt32(datatables[1].Rows[index]["XLOC Revise To Pendig"]);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void GetUniquePartNumber(string tablename)
        {
            string partnumber = string.Empty;
            List<Int32> indexlist = null;

            uniquepartnumberdic.Clear();

            try
            {
                if (tablename.Equals("FB Inventory"))
                {
                    for (int index = 0; index < datatables[0].Rows.Count; index++)
                    {
                        partnumber = datatables[0].Rows[index]["Usage P/N"].ToString();

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

                if (tablename.Equals("FB Variance GL&Recon"))
                {
                    for (int index = 0; index < datatables[1].Rows.Count; index++)
                    {
                        partnumber = datatables[1].Rows[index]["P/N"].ToString();

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
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        } 

        //private List<PartItem> GetUniquePartNumber(string tablename)
        //{
        //    bool flag = false;
        //    PartItem temppartitem = null;
        //    List<PartItem> partitemlist = new List<PartItem>();

        //    try
        //    {
        //        if (tablename.Equals("FB Inventory"))
        //        {
        //            for (int index = 0; index < datatables[0].Rows.Count; index++)
        //            {
        //                flag = false;

        //                for (int indey = 0; indey < partitemlist.Count; indey++)
        //                {
        //                    if (string.Compare(datatables[0].Rows[index]["Usage P/N"].ToString(), partitemlist[indey].PartNumber,
        //                        StringComparison.OrdinalIgnoreCase) == 0)
        //                    {
        //                        flag = true;
        //                        partitemlist[indey].Index.Add(index);
        //                        break;
        //                    }
        //                }

        //                if (!flag)
        //                {
        //                    temppartitem = new PartItem();
        //                    temppartitem.PartNumber = datatables[0].Rows[index]["Usage P/N"].ToString();
        //                    temppartitem.Index = new List<Int32>();
        //                    temppartitem.Index.Add(index);

        //                    partitemlist.Add(temppartitem);
        //                }
        //            }
        //        }

        //        if (tablename.Equals("FB Variance GL&Recon"))
        //        {
        //            for (int index = 0; index < datatables[1].Rows.Count; index++)
        //            {
        //                flag = false;

        //                for (int indey = 0; indey < partitemlist.Count; indey++)
        //                {
        //                    if (string.Compare(datatables[1].Rows[index]["P/N"].ToString(), partitemlist[indey].PartNumber,
        //                        StringComparison.OrdinalIgnoreCase) == 0)
        //                    {
        //                        flag = true;
        //                        partitemlist[indey].Index.Add(index);
        //                        break;
        //                    }
        //                }

        //                if (!flag)
        //                {
        //                    temppartitem = new PartItem();
        //                    temppartitem.PartNumber = datatables[0].Rows[index]["P/N"].ToString();
        //                    temppartitem.Index = new List<Int32>();
        //                    temppartitem.Index.Add(index);

        //                    partitemlist.Add(temppartitem);
        //                }
        //            }
        //        }                
        //    }
        //    catch (Exception ex)
        //    {
        //        MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
        //        throw;
        //    }

        //    return partitemlist;
        //}
    }
}

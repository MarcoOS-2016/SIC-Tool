using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using log4net;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight;

namespace SIC_Tool.Common
{
    public class ExcelFileUtility
    {
        private static readonly ILog log = LogManager.GetLogger(typeof (ExcelFileUtility));

        // Export report into a CSV file
        public static void ExportDataIntoExcelFile(string filename, DataTable datatable)
        {
            if (filename.Length != 0)
            {
                FileStream filestream = null;
                StreamWriter streamwriter = null;
                string stringline = string.Empty;

                try
                {
                    filestream = new FileStream(filename, FileMode.Append, FileAccess.Write);
                    streamwriter = new StreamWriter(filestream, Encoding.Unicode);

                    for (int i = 0; i < datatable.Columns.Count; i++)
                    {
                        stringline = stringline + datatable.Columns[i].ColumnName + Convert.ToChar(9);
                    }

                    streamwriter.WriteLine(stringline);
                    stringline = "";

                    for (int i = 0; i < datatable.Rows.Count; i++)
                    {
                        //stringline = stringline + (i + 1) + Convert.ToChar(9);
                        for (int j = 0; j < datatable.Columns.Count; j++)
                        {
                            //stringline = stringline + ((char)(9)).ToString() + datatable.Rows[i][j] + Convert.ToChar(9);
                            stringline = stringline + datatable.Rows[i][j] + Convert.ToChar(9);
                        }

                        streamwriter.WriteLine(stringline);
                        stringline = "";
                    }
                }
                catch (Exception ex)
                {
                    log.Info(ex.Message);
                    throw;
                }
                finally
                {
                    if (streamwriter != null)
                    {
                        streamwriter.Close();
                        streamwriter = null;
                    }

                    if (filestream != null)
                    {
                        filestream.Close();
                        filestream = null;
                    }
                    
                    GC.Collect();
                }
            }
        }

        // Retrive a sheet name from Excel file by number of sheet
        public static void ChangeColumnDataType(string fullfilename, string sheetname, List<string> columnnamelist,
            string datatype)
        {
            Application excel = null;
            Workbook workbook = null;
            Sheets worksheets = null;
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            //string sheetName = String.Empty;

            try
            {
                excel = new Application();

                excel.Visible = false;
                excel.DisplayAlerts = false;

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                workbook = excel.Workbooks.Open(fullfilename, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true,
                    false, 0, true, false, false);
                worksheets = workbook.Worksheets;

                foreach (Worksheet worksheet in worksheets)
                {
                    Range range = worksheet.UsedRange;

                    int colCount = range.Columns.Count;
                    int rowCount = range.Rows.Count;

                    if (worksheet.Name.Contains(sheetname))
                    {
                        foreach (string columnname in columnnamelist)
                        {
                            for (int index = 1; index <= colCount; index++)
                            {
                                if (string.Compare(worksheet.Cells[1, index].Value, columnname) == 0)
                                {
                                    for (int indey = 2; indey < rowCount - 1; indey++)
                                    {
                                        //worksheet.Cells[indey, index].NumberFormatLocal = datatype;
                                        worksheet.Cells[indey, index] = datatype + worksheet.Cells[indey, index].Value;
                                    }

                                    break;
                                }
                            }
                        }
                    }
                }

                //SetTitusClassification(ref workbook);

                workbook.Save();
            }

            catch (Exception ex)
            {
                log.Info(ex.Message);
                throw;
            }

            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                {
                    workbook.Close(true);
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        // Export report to an excel file
        public static void SaveExcelFile(string filename, DataTable datatable, bool iswithline)
        {
            Application excel = null;
            Workbook workbook = null;
            //Worksheet worksheet = null;

            try
            {
                excel = new Application();

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                excel.Application.Workbooks.Add(true);
                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Add();
                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

                // Write column name into Excel file
                int colIndex = 0;
                foreach (DataColumn col in datatable.Columns)
                {
                    colIndex++;
                    excel.Cells[1, colIndex] = col.ColumnName;
                }

                Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, colIndex]];
                if (iswithline)
                {
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                }

                // Write row data into Excel file
                int rowcount = datatable.Rows.Count;
                int colcount = datatable.Columns.Count;

                if (rowcount != 0 && colcount != 0)
                {
                    var dataarray = new object[rowcount, colcount];

                    for (int indey = 0; indey < rowcount; indey++)
                    {
                        for (int indez = 0; indez < colcount; indez++)
                        {
                            dataarray[indey, indez] = datatable.Rows[indey][indez];
                        }
                    }

                    range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[rowcount + 1, colcount]];
                    range.Value = dataarray;
                }

                if (iswithline)
                {
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                }

                worksheet.Cells.EntireColumn.AutoFit();
                SetTitusClassification(ref workbook);
                workbook.SaveAs(filename);
            }

            catch (Exception e)
            {
                log.Info(e.Message);
                throw;
            }

            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        // Export multiple reports to an excel file with multiple sheets
        public static void SaveExcelFileWithMultipleSheets(string filename, List<DataTable> datatablelist, bool iswithline)
        {
            Application excel = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            //Sheets sheets = null;
            Range range = null;

            try
            {
                excel = new Application();

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                excel.Visible = false;
                excel.DisplayAlerts = false;

                excel.Application.Workbooks.Add(true);
                workbook = excel.Workbooks.Add();
                //sheets = workbook.Worksheets;

                for (int index = 0; index < datatablelist.Count; index++)
                {
                    if (worksheet == null)
                    {
                        worksheet = (Worksheet) workbook.Worksheets[1];
                    }
                    else
                    {
                        worksheet = (Worksheet) workbook.Worksheets.Add(Type.Missing, worksheet, 1, Type.Missing);
                    }

                    worksheet.Name = datatablelist[index].TableName;

                    // Write column name into Excel file
                    int colIndex = 0;
                    foreach (DataColumn col in datatablelist[index].Columns)
                    {
                        colIndex++;
                        excel.Cells[1, colIndex] = col.ColumnName;
                    }

                    range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, colIndex]];

                    if (iswithline)
                    {
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlThin;
                    }

                    // Write row data into Excel file
                    int rowcount = datatablelist[index].Rows.Count;
                    int colcount = datatablelist[index].Columns.Count;

                    if (rowcount != 0 && colcount != 0)
                    {
                        var dataarray = new object[rowcount, colcount];

                        for (int indey = 0; indey < rowcount; indey++)
                        {
                            for (int indez = 0; indez < colcount; indez++)
                            {
                                dataarray[indey, indez] = datatablelist[index].Rows[indey][indez];
                            }
                        }

                        range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[rowcount + 1, colcount]];
                        range.Value = dataarray;
                    }

                    if (iswithline)
                    {
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlThin;
                    }

                    worksheet.Cells.EntireColumn.AutoFit();
                }

                SetTitusClassification(ref workbook);

                workbook.SaveAs(filename);
            }
            catch (Exception e)
            {
                log.Info(e.Message);
                throw;
            }

            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        // Export multiple reports to an excel file
        public static void SaveExcelFile(string filename, List<DataTable> datatablelist)
        {
            Application excel = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            //Sheets sheets = null;

            try
            {
                excel = new Application();

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                excel.Visible = false;
                excel.DisplayAlerts = false;

                excel.Application.Workbooks.Add(true);
                workbook = excel.Workbooks.Add();
                //sheets = workbook.Worksheets;

                for (int index = 0; index < datatablelist.Count; index++)
                {
                    if (worksheet == null)
                    {
                        worksheet = (Worksheet) workbook.Worksheets[1];
                    }
                    else
                    {
                        worksheet = (Worksheet) workbook.Worksheets.Add(Type.Missing, worksheet, 1, Type.Missing);
                    }

                    //worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[index + 1];
                    worksheet.Name = datatablelist[index].TableName;

                    int rowIndex = 1;
                    int colIndex = 0;

                    foreach (DataColumn col in datatablelist[index].Columns)
                    {
                        colIndex++;
                        excel.Cells[1, colIndex] = col.ColumnName;
                    }

                    foreach (DataRow row in datatablelist[index].Rows)
                    {
                        rowIndex++;
                        colIndex = 0;

                        string content = string.Empty;
                        foreach (DataColumn col in datatablelist[index].Columns)
                        {
                            colIndex++;
                            content = row[col.ColumnName].ToString();

                            if (col.ColumnName.ToUpper().Contains("DATE"))
                            {
                                excel.Cells[rowIndex, colIndex].NumberFormatLocal = "yyyy-mm-dd";
                                excel.Cells[rowIndex, colIndex] = content;
                            }
                            else if (col.ColumnName.ToUpper().Contains("SERVICE ORDER")
                                     || col.ColumnName.ToUpper().Contains("USAGE P/N")
                                     || col.ColumnName.ToUpper().Contains("DISPATCH P/N"))
                            {
                                excel.Cells[rowIndex, colIndex].NumberFormatLocal = "@"; //Set date type as text
                                excel.Cells[rowIndex, colIndex] = content;
                            }
                            else
                            {
                                excel.Cells[rowIndex, colIndex] = content;
                            }
                        }
                    }

                    //if (datatablelist[index].TableName.Equals("Variance_Detail_Items"))
                    //excel.Cells.Sort(excel.Cells.Columns[3], Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending); // Column[3] indicates to "Type" column.

                    worksheet.Cells.EntireColumn.AutoFit();
                }

                //workbook.SaveAs(String.Format("{0}_{1}.xls", filename, DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                SetTitusClassification(ref workbook);
                workbook.SaveAs(filename);
            }

            catch (Exception e)
            {
                log.Info(e.Message);
                throw;
            }

            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        //Save non-standard excel file format as standard one(.xlsx)
        public static string SaveAsStandardFileFormat(string fullfilename)
        {
            Application excel = null;
            Workbook workbook = null;

            try
            {
                string newfilename = string.Empty;
                excel = new Application();

                excel.Visible = false;
                excel.DisplayAlerts = false;

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                workbook = excel.Workbooks.Open(fullfilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);

                if (workbook.FileFormat != XlFileFormat.xlWorkbookDefault)
                {
                    MiscUtility.LogHistory(string.Format("Starting convert file format of excel report - {0}...",
                        fullfilename));

                    string filepath = Path.GetDirectoryName(fullfilename);
                    newfilename = Path.Combine(filepath, String.Format("{0}_{1}",
                        Path.GetFileNameWithoutExtension(fullfilename),
                        DateTime.Now.ToString("yyyyMMdd_HHmmss")));

                    SetTitusClassification(ref workbook);
                    workbook.SaveAs(newfilename, XlFileFormat.xlWorkbookDefault);

                    fullfilename = Path.ChangeExtension(newfilename, ".xlsx");
                        // Change the extension file name as Excel 2010's format

                    MiscUtility.LogHistory("Done!");
                }

                workbook.Close(true);
                excel.Quit();

                return fullfilename;
            }

            catch (Exception ex)
            {
                log.Info(ex.Message);
                throw;
            }

            finally
            {
                if (excel != null)
                {
                    Marshal.ReleaseComObject(excel);
                    excel = null;
                }

                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        //Save non-standard excel file format as standard one(.xlsx)
        public static string SaveAsStandardFileFormat(string fullfilename, string sheet1name)
        {
            Application excel = null;
            Workbook workbook = null;

            try
            {
                string newfilename = string.Empty;
                excel = new Application();

                excel.Visible = false;
                excel.DisplayAlerts = false;

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                workbook = excel.Workbooks.Open(fullfilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);

                if (sheet1name.Length != 0)
                    ((Worksheet) workbook.Worksheets[1]).Name = sheet1name;

                if (workbook.FileFormat != XlFileFormat.xlWorkbookDefault)
                {
                    FileUtility.SaveFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "History.txt"),
                        string.Format("[{0}] - Starting convert file format of excel report - {1}...", DateTime.Now,
                            fullfilename));

                    string filepath = Path.GetDirectoryName(fullfilename);
                    newfilename = Path.Combine(filepath, String.Format("{0}_{1}",
                        Path.GetFileNameWithoutExtension(fullfilename),
                        DateTime.Now.ToString("yyyyMMdd_HHmmss")));

                    SetTitusClassification(ref workbook);

                    workbook.SaveAs(newfilename, XlFileFormat.xlWorkbookDefault);

                    fullfilename = Path.ChangeExtension(newfilename, ".xlsx");
                        // Change the extension file name as Excel 2010's format

                    FileUtility.SaveFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "History.txt"),
                        string.Format("[{0}] - Done!", DateTime.Now));
                }

                workbook.Close(true);
                excel.Quit();

                return fullfilename;
            }

            catch (Exception ex)
            {
                log.Info(ex.Message);
                throw;
            }

            finally
            {
                if (excel != null)
                {
                    Marshal.ReleaseComObject(excel);
                    excel = null;
                }

                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        // Add some properties for Titus Classification into Excel file.
        public static void SetTitusClassification(ref Workbook workBook)
        {
            SetDocumentProperty(ref workBook, "DellClassification", "Internal Use");
            SetDocumentProperty(ref workBook, "TitusReset", "Reset");
        }

        // Setup a customer property for excel file.
        public static void SetDocumentProperty(ref Workbook workBook,
            string propertyName, string propertyValue)
        {
            dynamic oDocCustomProps = workBook.CustomDocumentProperties;
            Type typeDocCustomProps = oDocCustomProps.GetType();

            dynamic[] oArgs = {propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue};

            typeDocCustomProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null,
                oDocCustomProps, oArgs);
        }

        public static dynamic GetDocumentProperty(ref Workbook workBook,
            string propertyName, MsoDocProperties type)
        {
            dynamic returnVal = null;

            dynamic oDocCustomProps = workBook.CustomDocumentProperties;
            Type typeDocCustomProps = oDocCustomProps.GetType();

            dynamic returned = typeDocCustomProps.InvokeMember("Item",
                BindingFlags.Default |
                BindingFlags.GetProperty, null,
                oDocCustomProps, new object[] {propertyName});

            Type typeDocAuthorProp = returned.GetType();
            returnVal = typeDocAuthorProp.InvokeMember("Value",
                BindingFlags.Default |
                BindingFlags.GetProperty,
                null, returned,
                new object[] {}).ToString();

            return returnVal;
        }

        #region ----- Append rows into an existing Excel file ------
        public static void AppendDateToExcelFile(string fullfilename, string sheetname, DataTable datatable)
        {
            Workbook workbook = null;
            Sheets worksheets = null;
            Worksheet worksheet = null;
            Application excel = null;

            try
            {
                excel = new Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workbook = excel.Workbooks.Open(fullfilename, 0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                worksheets = workbook.Worksheets;

                worksheet = (Worksheet) worksheets.get_Item(sheetname);
                Range range = worksheet.UsedRange;

                int colCount = range.Columns.Count;
                int rowCount = range.Rows.Count;

                for (int colindex = 0; colindex < datatable.Columns.Count; colindex++)
                {
                    for (int rowindex = 0; rowindex < datatable.Rows.Count; rowindex++)
                    {
                        worksheet.Cells[rowCount + rowindex + 1, colindex + 1] = datatable.Rows[rowindex][colindex];
                    }
                }

                workbook.SaveAs(fullfilename, XlFileFormat.xlWorkbookNormal,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive,
                    Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);

            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                throw;
            }
            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

                if (workbook != null)
                {
                    workbook.Close(true);
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        #endregion
    }
}
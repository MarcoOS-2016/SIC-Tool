using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using SIC_Tool.Common;

namespace SIC_Tool.Business
{
    public class ExcelManager : IDisposable
    {
        private Microsoft.Office.Interop.Excel.Application excelinstance = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel.Sheets worksheets = null;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        private Microsoft.Office.Interop.Excel.Range range = null;

        public Microsoft.Office.Interop.Excel.Application ExcelInstance
        {
            get { return excelinstance; }
        }

        public Microsoft.Office.Interop.Excel.Workbook WorkBook
        {
            get { return workbook; }
            set { workbook = value; }
        }

        public Microsoft.Office.Interop.Excel.Sheets WorkSheets
        {
            get { return worksheets; }
            set { worksheets = value; }
        }

        public Microsoft.Office.Interop.Excel.Worksheet WorkSheet
        {
            get { return worksheet; }
            set { worksheet = value; }
        }

        public Microsoft.Office.Interop.Excel.Range Range
        {
            get { return range; }
            set { range = value; }
        }

        public ExcelManager()
        {
            try
            {
                if (excelinstance == null)
                {
                    excelinstance = new Microsoft.Office.Interop.Excel.Application();
                    if (excelinstance == null)
                        throw new Exception("There is not an Excel application on your computer!");
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
        }

        public ExcelManager(string fullfilename)
        {
            try
            {
                if (excelinstance == null)
                {
                    excelinstance = new Microsoft.Office.Interop.Excel.Application();
                    if (excelinstance == null)
                        throw new Exception("There is not an Excel application on your computer!");

                    excelinstance.Visible = false;
                    excelinstance.DisplayAlerts = false;

                    this.workbook = excelinstance.Workbooks.Open(fullfilename, 0, false, 5, "", "", false,
                        Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
        }

        public void OpenWorkBook(string fullfilename)
        {
            try
            {
                excelinstance.Visible = false;
                excelinstance.DisplayAlerts = false;

                this.workbook = excelinstance.Workbooks.Open(fullfilename, 0, false, 5, "", "", false,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //return workbook;
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
        }

        public List<string> GetSheetNameList()
        {
            List<string> sheetnamelist = new List<string>();

            for (int index = 0; index < excelinstance.Workbooks[1].Worksheets.Count; index++)
            {
                sheetnamelist.Add(excelinstance.Workbooks[1].Worksheets[index].Name);
            }

            return sheetnamelist;
        }

        public Worksheet GetSheetByName(string sheetname)
        {
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelinstance.Worksheets[sheetname];
            return worksheet;
        }

        public Worksheet GetSheetByNumber(int sheetnumber)
        {
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelinstance.Worksheets[sheetnumber];
            return worksheet;
        }

        public void DeleteRow(ref Worksheet worksheet, string sheetname, int rowindex)
        {
            excelinstance.Worksheets[sheetname].Rows[rowindex].Delete(true);
        }

        public void InsertColumn(string columnname, List<string> rawdatalist)
        {

        }

        public void SaveWorkBook()
        {
            try
            {
                if (workbook != null)
                {
                    workbook.Save();
                    //ExcelFileUtility.SetTitusClassification(ref workbook);
                    workbook.Close(true);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
        }

        public void SaveWorkBook(string fullfilename)
        {
            try
            {
                if (workbook != null)
                {
                    ExcelFileUtility.SetTitusClassification(ref workbook);
                    workbook.SaveAs(fullfilename);
                    workbook.Close(true);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
        }

        ~ExcelManager()
        {              
            Dispose();
        }

        public void Dispose()
        {
            //Dispose(true);
            //GC.SuppressFinalize(this);

            if (this.worksheet != null)
                ReleaseObject(worksheet);

            if (this.worksheets != null)
                ReleaseObject(worksheets);

            if (this.workbook != null)
                ReleaseObject(workbook);

            if (this.excelinstance != null)
            {
                this.excelinstance.Quit();
                ReleaseObject(excelinstance);
            }

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        private void ReleaseObject(object objectinstance)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objectinstance);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(ex.Message);
                throw ex;
            }
            finally
            {
                objectinstance = null;
            }
        }

        //protected virtual void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {
        //        // Free managed objects.
        //    }
        //    // Free unmanaged objects.

        //    if (this.excelinstance != null)
        //    {
        //        this.excelinstance.Quit();
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelinstance);
        //        this.excelinstance = null;
        //    }

        //    GC.SuppressFinalize(this);
        //}
    }
}

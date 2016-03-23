using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.DataAccess
{
    public class ExcelAccessDAO : ExcelFileAccess
    {
        public ExcelAccessDAO()
        {
        }

        public ExcelAccessDAO(string filename)
            : base(filename)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public ExcelAccessDAO(string filename, bool isfield)
            : base(filename, isfield)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public DataTable GetExcelSheetName()
        {            
            DataTable schematable = base.connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
            return schematable;
        }

        public DataSet GetServiceTagData(string sheetname)
        {
            string newsheetname = sheetname.Replace("'", "");
            //string sqlString = String.Format("SELECT DISTINCT [Service tag] FROM [{0}] WHERE [Dell sales order number] LIKE '100%';", newsheetname);
            string sqlString = String.Format("SELECT DISTINCT [Service tag] FROM [{0}];", newsheetname);
            return this.ExecuteQuery(sqlString);
        }

        public DataSet ReadExcelFile(string sheetname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT * FROM [{0}];", sheetname);
            else
                sqlString = String.Format("SELECT * FROM [{0}$];", sheetname);

            return this.ExecuteQuery(sqlString);
        }

        public DataSet ReadExcelFile(string sheetname, string fieldname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT {0} FROM [{1}];", fieldname, sheetname);
            else
                sqlString = String.Format("SELECT {0} FROM [{1}$];", fieldname, sheetname);

            return this.ExecuteQuery(sqlString);
        }
                
        //public DataSet GetPartUnitCost(string sheetname)
        //{
        //    string sqlString = string.Empty;

        //    if (sheetname.Contains("$"))
        //        sqlString = String.Format("SELECT DISTINCT ITEM, PARTCOST FROM [{0}]", sheetname);
        //    else
        //        sqlString = string.Format("SELECT DISTINCT ITEM, PARTCOST FROM [{0}$]", sheetname);

        //    return this.ExecuteQuery(sqlString);
        //}

        public DataSet GetPartData(string sheetname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT DISTINCT Item, Description, CC, PartCost FROM [{0}]", sheetname);
            else
                sqlString = String.Format("SELECT DISTINCT Item, Description, CC, PartCost FROM [{0}$]", sheetname);

            return this.ExecuteQuery(sqlString);
        }

        public DataSet GetInventoryEndingBalanceData(string sheetname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT * FROM [{0}] WHERE BIN <> 'GICBC'", sheetname);
            else
                sqlString = string.Format("SELECT * FROM [{0}$] WHERE BIN <> 'GICBC'", sheetname);

            return this.ExecuteQuery(sqlString);
        }
    }
}

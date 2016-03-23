using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using SIC_Tool.Common;

namespace SIC_Tool.DataAccess
{
    public class OracleAccessDAO : OracleDataBaseAccess
    {
        private string connectionstring = string.Empty;

        public OracleAccessDAO(string connectionstring)
            : base(connectionstring)
        {
            this.connectionstring = connectionstring;

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

        public DataSet FetchDataFromDataBase(string sqlString)
        {
            return this.ExecuteQuery(sqlString);
        }

        public DataSet GetPartDetail(string ccn)
        {
            DataSet ds = null;
            string sqlstring = string.Empty;

            if (ccn.Equals("BRH"))      //BRH doesn't have PAU field in Glovia database
            {
                sqlstring =
                    string.Format("SELECT DISTINCT D.CCN, D.ITEM, C.DESCRIPTION, C.COMMODITY, B.ISSUE_CODE, B.BOX_CODE, B.Bulk_Expensed, B.Consigned, B.COA, B.Return_On_ASN FROM glovia_prog40.C_ITEM A, glovia_prog40.C_ITMCCN B, glovia_prog40.ITEM C, glovia_prog40.ITEM_DET D, glovia_prog40.COMMOD E WHERE D.ITEM = C.ITEM AND C.ITEM = A.ITEM AND A.ITEM = B.ITEM AND B.Item = D.Item AND B.CCN = D.CCN AND D.REVISION = C.REVISION AND C.REVISION = A.REVISION AND A.REVISION = B.REVISION AND B.REVISION = D.REVISION AND C.COMMODITY = E.COMMODITY AND D.CCN = '{0}'", ccn);
            }
            else
            {
                sqlstring =
                    string.Format("SELECT DISTINCT D.CCN, D.ITEM, C.DESCRIPTION, C.COMMODITY, B.ISSUE_CODE, B.BOX_CODE, B.Bulk_Expensed, B.Consigned, B.COA, B.Return_On_ASN, B.PAU FROM glovia_prog40.C_ITEM A, glovia_prog40.C_ITMCCN B, glovia_prog40.ITEM C, glovia_prog40.ITEM_DET D, glovia_prog40.COMMOD E WHERE D.ITEM = C.ITEM AND C.ITEM = A.ITEM AND A.ITEM = B.ITEM AND B.Item = D.Item AND B.CCN = D.CCN AND D.REVISION = C.REVISION AND C.REVISION = A.REVISION AND A.REVISION = B.REVISION AND B.REVISION = D.REVISION AND C.COMMODITY = E.COMMODITY AND D.CCN = '{0}'", ccn);
            }
            ds = this.ExecuteQuery(sqlstring);

            return ds;
        }

        public DataSet GetPartDetailByPartNumber(string ccn, string partnumberlist)
        {
            DataSet ds = new DataSet();
            string sqlstring = string.Empty;

            if (ccn.Equals("BRH"))
            {
                sqlstring =
                    string.Format("SELECT DISTINCT B.CCN, B.ITEM, C.DESCRIPTION, C.COMMODITY, B.ISSUE_CODE, B.BOX_CODE, B.Bulk_Expensed, B.Consigned, B.COA, B.Return_On_ASN FROM glovia_prog40.C_ITEM A, glovia_prog40.C_ITMCCN B, glovia_prog40.ITEM C, glovia_prog40.COMMOD E WHERE C.ITEM = A.ITEM AND A.ITEM = B.ITEM AND C.REVISION = A.REVISION AND C.COMMODITY = E.COMMODITY AND B.CCN = '{0}' AND C.ITEM IN ( {1} )", ccn, partnumberlist);
            }
            else
            {
                sqlstring =
                    string.Format("SELECT DISTINCT B.CCN, B.ITEM, C.DESCRIPTION, C.COMMODITY, B.ISSUE_CODE, B.BOX_CODE, B.Bulk_Expensed, B.Consigned, B.COA, B.Return_On_ASN, B.PAU FROM glovia_prog40.C_ITEM A, glovia_prog40.C_ITMCCN B, glovia_prog40.ITEM C, glovia_prog40.COMMOD E WHERE C.ITEM = A.ITEM AND A.ITEM = B.ITEM AND C.REVISION = A.REVISION AND C.COMMODITY = E.COMMODITY AND B.CCN = '{0}' AND C.ITEM IN ( {1} )", ccn, partnumberlist);
            }

            //for (int count = 0; count < partnumberlist.Count; count++)
            //{
            //    if (count > 0 && (count % 1000 == 0))
            //    {
            //        sqlstring.Remove(sqlstring.Length - 1, 1);
            //        sqlstring.Append(") OR C.ITEM IN (");
            //    }

            //    sqlstring.Append("'");
            //    sqlstring.Append(partnumberlist[count]);
            //    sqlstring.Append("',");
            //}

            //sqlstring.Remove(sqlstring.Length - 1, 1);
            //sqlstring.Append(")");
            ds = this.ExecuteQuery(sqlstring);

            FileUtility.SaveFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sql.txt"),
                string.Format("{0} ----- The SQL Script of Getting Part Detail: {1}", DateTime.Now.ToString(), sqlstring));

            return ds;
        }

        public DataSet GetFGATransactionRecord(List<string> servicetaglist)
        {
            DataSet ds = new DataSet();

            string sqlstring = "SELECT a.service_tag, a.item, b.commodity, a.item_qty, a.transaction_type, a.tot_item_cost, a.matl_move_from_loc, a.matl_move_to_loc, a.relief_from_loc, a.consigned_mat, a.transaction_date FROM dell_inv_svctag_dtl a inner join item b on a.item = b.item WHERE a.transaction_type in ('POUPDATE','XFER') AND (service_tag in ";

            StringBuilder sb = new StringBuilder(sqlstring);
            for (int count = 0; count < servicetaglist.Count; count++)
            {
                sb.Append(string.Format("({0}) ", servicetaglist[count]));
                sb.Append("OR service_tag IN ");
            }

            sb.Remove(sb.Length - "OR service_tag IN ".Length, "OR service_tag IN ".Length);
            sb.Append(")");

            FileUtility.SaveFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sql.txt"),
                string.Format("{0} ----- The SQL Script of Getting FGA Transaction History: {1}", DateTime.Now.ToString(), sb.ToString()));

            ds = this.ExecuteQuery(sb.ToString());

            return ds;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using log4net;

namespace SIC_Tool.DataAccess
{
    public class OracleDataBaseAccess : IDisposable
    {
        private static ILog log = LogManager.GetLogger(typeof(OracleDataBaseAccess));
        //private string connectionString = String.Empty;
        protected int affectrow = 0;
        protected OracleConnection connection = null;

        public int AffectRow
        {
            get
            {
                return this.affectrow;
            }
        }

        public OracleConnection Connection
        {
            get
            {
                return this.connection;
            }
            set
            {
                this.connection = value;
            }
        }

        //public OracleDataBaseAccess()
        //{
        //    connection = new OracleConnection();
        //    connection.ConnectionString = ConfigurationManager.AppSettings["OracleConnection"];

        //    try
        //    {
        //        connection.Open();
        //    }
        //    catch (Exception e)
        //    {
        //        log.Error(e.Message);
        //        throw e;
        //    }
        //}

        public OracleDataBaseAccess(string connectionstring)
        {
            connection = new OracleConnection();
            connection.ConnectionString = connectionstring;

            try
            {
                connection.Open();
            }

            catch (Exception e)
            {
                log.Error(e.Message);
                throw e;
            }
        }

        public int ExecuteNonQuery(string sql)
        {
            OracleCommand cmd = new OracleCommand(sql, this.connection);
            cmd.CommandType = CommandType.Text;

            try
            {
                if (this.connection.State != ConnectionState.Open)
                {
                    cmd.Connection.Open();
                }

                this.affectrow = cmd.ExecuteNonQuery();

            }

            catch (Exception e)
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                    connection.Dispose();
                }

                log.Error(string.Format(sql, e.Message));
                throw e;
            }

            finally
            {
                cmd.Dispose();
            }

            return this.affectrow;
        }

        public OracleDataReader ExecuteReader(string sql)
        {
            OracleCommand cmd = new OracleCommand(sql, this.connection);
            cmd.CommandType = CommandType.Text;

            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    cmd.Connection.Open();
                }

                OracleDataReader reader = cmd.ExecuteReader();

                return reader;
            }

            catch (Exception e)
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                    connection.Dispose();
                }

                log.Error(string.Format(sql, e.Message));
                throw e;
            }

            finally
            {
                cmd.Dispose();
            }
        }

        public DataSet ExecuteQuery(string sql)
        {
            DataSet ds = null;
            OracleDataAdapter adapter = null;

            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                adapter = new OracleDataAdapter(sql, connection);
                ds = new DataSet();
                adapter.Fill(ds);
            }

            catch (Exception e)
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                    connection.Dispose();
                }

                log.Error(string.Format(sql, e.Message));
                throw e;
            }

            finally
            {
                adapter.Dispose();
            }

            return ds;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed objects.
            }
            // Free unmanaged objects.

            if (this.connection != null)
            {
                this.connection.Close();
                this.connection.Dispose();
                this.connection = null;
            }

            // Set large fields to null.
        }

        ~OracleDataBaseAccess()
        {
            if (this.connection != null && this.connection.State == ConnectionState.Open)
            {
                log.Error("Connection is supposed to be closed by clients instead of waiting for GC.");
            }

            Dispose(false);
        }
    }
}

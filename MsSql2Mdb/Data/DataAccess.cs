using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsSql2Mdb.Data
{
    public class DataAccess
    {
        static string _dbConnectionStr = @"";

        public static string DbConnectionStr
        {
            get
            {
                if (string.IsNullOrEmpty(_dbConnectionStr))
                {
                    _dbConnectionStr = ConfigurationManager.ConnectionStrings["DbConnectionStr"].ConnectionString;
                }
                return _dbConnectionStr;
            }
        }
        public DataAccess() { }

        SqlConnection _cnn = null;
        SqlConnection Cnn
        {
            get
            {
                if(_cnn==null)
                {
                    _cnn = new SqlConnection(DbConnectionStr);
                }
                return _cnn;
            }
        }
        //使用 TableName 取得 DataTable
        public DataTable GetTableData(string TableName, string SchemaName)
        {
            try
            {
                SqlCommand cmd = new SqlCommand(string.Format("select * from [{1}].[{0}]", TableName, SchemaName), Cnn);
                DataSet ds = new DataSet();
                SqlDataAdapter SqlDa = new SqlDataAdapter(cmd);
                SqlDa.Fill(ds);
                return ds.Tables[0];
            }
            finally
            {
                Cnn.Close();
            }
        }
        //使用 DataColumnCollection & TableName 取得 SQL Insert Statement.
        public string GetInsertCommandByDataColumns(DataColumnCollection columns, string TableName, DataRow row)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("INSERT INTO [{0}]", TableName));

            sb.Append("(");
            for (int i = 0; i < columns.Count; i++)
            {
                if (i < columns.Count - 1)
                {
                    sb.AppendFormat("[{0}], ", columns[i].ColumnName);
                }
                else
                {
                    sb.AppendFormat("[{0}]", columns[i].ColumnName);
                }
            }
            sb.AppendLine(")");

            sb.Append("values(");
            for (int i = 0; i < columns.Count; i++)
            {
                string p = GetParameterValue(columns[i].DataType, row[columns[i].ColumnName]);

                if (i < columns.Count - 1)
                {
                    if (row[columns[i].ColumnName] != DBNull.Value)
                    {
                        if (columns[i].DataType == typeof(byte[]))
                        {
                            sb.Append(string.Format("{0},", "null"));
                        }
                        else
                        {
                            sb.Append(string.Format("'{0}',", p.Replace("\r\n", "").Replace("'", "")));
                        }
                    }
                    else
                    {
                        sb.Append(string.IsNullOrEmpty(p) ? string.Format("{0},", "null") : string.Format("'{0}',", "null"));
                    }
                }
                else
                {
                    if (row[columns[i].ColumnName] != DBNull.Value)
                    {
                        if (columns[i].DataType == typeof(byte[]))
                        {
                            sb.Append(string.Format("{0}", "null"));
                        }
                        else
                        {
                            sb.Append(string.IsNullOrEmpty(p) ? string.Format("'{0}'", p) : string.Format("'{0}'", p.Replace("\r\n", "").Replace("'", "")));
                        }
                    }
                    else
                    {
                        sb.Append(string.IsNullOrEmpty(p) ? string.Format("{0}", "null") : string.Format("'{0}'", p.Replace("\r\n", "").Replace("'", "")));
                    }
                }
            }
            sb.Append(")");
            return sb.ToString();
        }
        //
        public OleDbParameter GetOleDbParameterByDataColumn(DataColumn column, DataRow row)
        {
            OleDbParameter SqlParam = null;
            switch (column.DataType.ToString())
            {
                case "System.String":
                case "System.Int32":
                case "System.DateTime":
                case "System.Decimal":
                case "System.Guid":
                case "System.Boolean":
                    SqlParam = new OleDbParameter(column.ColumnName, row[column.ColumnName]);
                    break;
                case "System.Byte[]":
                    SqlParam = new OleDbParameter(column.ColumnName, DBNull.Value);
                    break;
                default:
                    SqlParam = new OleDbParameter(column.ColumnName, DBNull.Value);
                    break;
            }
            return SqlParam;
        }
        public string GetParameterValue(Type ColumnType, object Value)
        {
            switch (ColumnType.ToString())
            {
                case "System.String":
                case "System.Int32":
                case "System.Int16":
                case "System.Int64":
                case "System.Decimal":
                case "System.Guid":

                    if (Value != DBNull.Value)
                        return Value.ToString().TrimEnd();
                    else
                        return "";
                //break;
                case "System.Byte[]":
                    return "";
                //break;
                case "System.DateTime":
                    if (Value != DBNull.Value)
                        return Convert.ToDateTime(Value).ToString("yyyy/MM/dd");
                    else
                        return "";
                case "System.Single":
                    return Value.ToString();
                    break;
                case "System.Boolean":
                    return Convert.ToInt32(Value).ToString();
                    break;
                default:
                    return "";
            }
        }
    }
}

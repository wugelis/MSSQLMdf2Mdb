using ADOX;
using MsSql2Mdb.Data.Schema;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsSql2Mdb.Data
{
    /// <summary>
    /// 建立 Access Mdb 資料庫檔案.
    /// </summary>
    public class MdbHelper
    {
        public static string OleDbConnectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}; Jet OLEDB:Engine Type=5";
        public static string OleDbConnectionStr2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};";
        //
        private static ADOX.DataTypeEnum GetMdbDataTypeByMSSQL(string TypeString)
        {
            switch (TypeString)
            {
                case "int":
                case "bigint":
                case "smallint":
                case "tinyint":
                case "bit":
                    return ADOX.DataTypeEnum.adInteger;
                case "nvarchar":
                case "uniqueidentifier":
                    return ADOX.DataTypeEnum.adLongVarWChar;
                    //return ADOX.DataTypeEnum.adBoolean;
                case "nchar":
                    return ADOX.DataTypeEnum.adVarWChar;
                case "datetime":
                    return ADOX.DataTypeEnum.adDate;
                case "money":
                    return ADOX.DataTypeEnum.adCurrency;
                    //return ADOX.DataTypeEnum.adSmallInt;
                case "ntext":
                    return ADOX.DataTypeEnum.adLongVarWChar;
                default:
                    return ADOX.DataTypeEnum.adVarWChar;
            }
        }
        /// <summary>
        /// 建立 Access 資料庫 MDB
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool CreateNewAccessDatabase(string fileName)
        {
            bool result = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table table = null;
            SchemaField traceSchema = null;

            //建立舊版的 ADODB CConnection 連線
            ADODB.Connection con = null;

            try
            {
                //建立 Access Mdb 資料庫.
                cat.Create(string.Format(OleDbConnectionStr, fileName));

                SchemaData Schema = new SchemaData();

                foreach (var tbName in Schema.GetTableSchemaData())
                {
                    var tbData = Schema.GetSchema().AsEnumerable().Where(c => c.TABLE_NAME == tbName.TABLE_NAME);

                    table = new ADOX.Table();
                    table.Name = tbName.TABLE_NAME;

                    foreach (var field in tbData)
                    {
                        traceSchema = field; //由於 COM 無法在IDE中監視內容值，加上此變數 Trace and Debug 用

                        if (field.CHARACTER_MAXIMUM_LENGTH.HasValue && (field.COLUMN_DATA_TYPE == "nvarchar" || field.COLUMN_DATA_TYPE == "nchar" || field.COLUMN_DATA_TYPE == "ntext" || field.COLUMN_DATA_TYPE == "varchar" || field.COLUMN_DATA_TYPE == "char"))
                        {
                            table.Columns.Append(field.COLUMN_NAME, GetMdbDataTypeByMSSQL(field.COLUMN_DATA_TYPE), (field.CHARACTER_MAXIMUM_LENGTH.Value > 255 || field.CHARACTER_MAXIMUM_LENGTH < 0) ? 255 : field.CHARACTER_MAXIMUM_LENGTH.Value);
                        }
                        else
                        {
                            table.Columns.Append(field.COLUMN_NAME, GetMdbDataTypeByMSSQL(field.COLUMN_DATA_TYPE));
                        }

                        if (field.IS_NULLABLE == "YES"
                            || field.COLUMN_DATA_TYPE == "binary"
                            || field.COLUMN_DATA_TYPE == "image"
                            || field.COLUMN_DATA_TYPE == "varbinary"
                            || field.COLUMN_DATA_TYPE == "uniqueidentifier") //從 GetSchema 的資訊中判斷該欄位是否為 NULL 並將該欄位設為允許 NULL (若欄位型態為 binary,image,varbinary 的話也將該欄位設為允許 NULL)
                        {
                            table.Columns[table.Columns.Count - 1].Attributes = ColumnAttributesEnum.adColNullable; //將目前的 欄位 設定為 IS NULL (允許空值)
                        }
                    }
                    cat.Tables.Append(table);
                }
                con = cat.ActiveConnection as ADODB.Connection;

                result = true;
            }
            catch (Exception ex)
            {
                Program.ShowError(ex.Message);
                result = false;
            }
            finally
            {
                if (con != null)
                    con.Close();
                cat = null;
            }
            return result;
        }
        /// <summary>
        /// 將資料一筆一筆 Insert 到 MDB 中.
        /// </summary>
        /// <param name="AccessMdb"></param>
        public static void TransferData2NewAccessDatabase(string AccessMdb)
        {
            SchemaData schema = new SchemaData();
            DataAccess dal = new DataAccess();
            var Tables = schema.GetTableSchemaData();

            OleDbConnection Conn = new OleDbConnection(string.Format(MdbHelper.OleDbConnectionStr2, AccessMdb));
            OleDbCommand Cmd = new OleDbCommand("", Conn);

            try
            {
                Conn.Open();

                #region 測試用
                //		DataTable dt = dal.GetTableData("Products");
                //							
                //		foreach(DataRow row in dt.Rows)
                //		{
                //			Cmd.CommandText = dal.GetInsertCommandByDataColumns(dt.Columns, "Products", row);
                //			Cmd.CommandText.Dump();
                //			
                //			Cmd.ExecuteNonQuery();
                //		}
                #endregion

                foreach (var schemaField in Tables)
                {
                    DataTable dt = dal.GetTableData(schemaField.TABLE_NAME, schemaField.SCHEMA_NAME);

                    foreach (DataRow row in dt.Rows)
                    {
                        Cmd.CommandText = dal.GetInsertCommandByDataColumns(dt.Columns, schemaField.TABLE_NAME, row);
                        //Console.WriteLine(Cmd.CommandText);

                        Cmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                Conn.Close();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsSql2Mdb.Data.Schema
{
    /// <summary>
    /// 
    /// </summary>
    public class SchemaData
    {
        SqlConnection cnn = new SqlConnection(DataAccess.DbConnectionStr);

        public IEnumerable<SchemaField> GetInitialCatalogData()
        {
            return new List<SchemaField>();
        }
        public DataTable GetBySchemaName(string SchemaName)
        {
            try
            {
                cnn.Open();
                return cnn.GetSchema();
            }
            finally
            {
                cnn.Close();
            }
        }
        public IEnumerable<SchemaField> GetTableSchemaData()
        {
            try
            {
                cnn.Open();
                DataTable dtCnnSchema = cnn.GetSchema("Tables");

                var result = from schema in dtCnnSchema.AsEnumerable()
                             where (string)schema["TABLE_TYPE"] == "BASE TABLE"
                             select new SchemaField
                             {
                                 DATABASE_NAME = (string)schema["TABLE_CATALOG"],
                                 SCHEMA_NAME = (string)schema["TABLE_SCHEMA"],
                                 TABLE_NAME = (string)schema["TABLE_NAME"],
                                 TABLE_TYPE = (string)schema["TABLE_TYPE"]
                             };
                return result.OrderBy(c => c.TABLE_NAME);
            }
            finally
            {
                cnn.Close();
            }
        }
        public IEnumerable<SchemaField> GetSchema()
        {
            try
            {
                cnn.Open();
                DataTable dtCnnSchema = cnn.GetSchema("Columns");
                //dtCnnSchema.AsEnumerable().Where(c => (string)c["TABLE_NAME"]=="Orders").Select(c => c). Dump();
                var result = from schema in dtCnnSchema.AsEnumerable()
                             select new SchemaField
                             {
                                 DATABASE_NAME = (string)schema["TABLE_CATALOG"],
                                 SCHEMA_NAME = (string)schema["TABLE_SCHEMA"],
                                 TABLE_NAME = (string)schema["TABLE_NAME"],
                                 COLUMN_NAME = (string)schema["COLUMN_NAME"],
                                 COLUMN_DATA_TYPE = (string)schema["DATA_TYPE"],
                                 IS_NULLABLE = (string)schema["IS_NULLABLE"],
                                 CHARACTER_MAXIMUM_LENGTH = schema["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value ? (int?)schema["CHARACTER_MAXIMUM_LENGTH"] : null
                             };
                return result.OrderBy(c => c.TABLE_NAME);
            }
            finally
            {
                cnn.Close();
            }
        }
        public DataTable GetSchemaOriginal()
        {
            try
            {
                cnn.Open();
                DataTable dtCnnSchema = cnn.GetSchema();
                return dtCnnSchema;
            }
            finally
            {
                cnn.Close();
            }
        }
        public DataTable GetSchemaOriginal(string collectionName)
        {
            try
            {
                if (string.IsNullOrEmpty(collectionName))
                    return GetSchemaOriginal();

                cnn.Open();
                DataTable dtCnnSchema = cnn.GetSchema(collectionName);
                return dtCnnSchema;
            }
            finally
            {
                cnn.Close();
            }
        }
    }
}

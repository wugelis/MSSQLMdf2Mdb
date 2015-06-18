using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsSql2Mdb.Data.Schema
{
    public class SchemaField
    {
        public string DATABASE_NAME { get; set; }
        public string SCHEMA_NAME { get; set; }
        public string TABLE_NAME { get; set; }
        public string TABLE_TYPE { get; set; }
        public string COLUMN_NAME { get; set; }
        public string COLUMN_DATA_TYPE { get; set; }
        public string IS_NULLABLE { get; set; }
        public int? CHARACTER_MAXIMUM_LENGTH { get; set; }
    }
}

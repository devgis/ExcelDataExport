using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace DataExport
{
    public class ExcelImport
    {
        /// <summary>
        /// Excel 版本
        /// </summary>
        public enum ExcelType
        {
            Excel2003, Excel2007
        }
        public enum IMEXType
        {
            ExportMode = 0, ImportMode = 1, LinkedMode = 2
        }
        public static string GetExcelFirstTableName(string excelPath, ExcelType eType)
        {
            string connectstring = GetExcelConnectstring(excelPath, true, eType);
            return GetExcelFirstTableName(connectstring);
        }
        public static string GetExcelFirstTableName(string connectstring)
        {
            using (OleDbConnection conn = new OleDbConnection(connectstring))
            {
                return GetExcelFirstTableName(conn);
            }
        }
        public static string GetExcelFirstTableName(OleDbConnection connection)
        {
            string tableName = string.Empty;

            if (connection.State == ConnectionState.Closed)
                connection.Open();

            DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt != null && dt.Rows.Count > 0)
            {
                tableName = ConvertTo<string>(dt.Rows[0][2]);
            }

            return tableName;
        }
        public static string GetExcelConnectstring(string excelPath, bool header, ExcelType eType)
        {
            return GetExcelConnectstring(excelPath, header, eType, IMEXType.ImportMode);
        }
        public static string GetExcelConnectstring(string excelPath, bool header, ExcelType eType, IMEXType imex)
        {
            if (!System.IO.File.Exists(excelPath)) ///
                throw new FileNotFoundException("Excel路径不存在!");

            string connectstring = string.Empty;

            string hdr = "NO";
            if (header)
                hdr = "YES";

            if (eType == ExcelType.Excel2003)
                connectstring = "Provider=Microsoft.Jet.OleDb.4.0; data source=" + excelPath + ";Extended Properties='Excel 8.0; HDR=" + hdr + "; IMEX=" + imex.GetHashCode() + "'";
            else
                connectstring = "Provider=Microsoft.ACE.OLEDB.12.0; data source=" + excelPath + ";Extended Properties='Excel 12.0 Xml; HDR=" + hdr + "; IMEX=" + imex.GetHashCode() + "'";

            return connectstring;
        }
        public static List<string> GetExcelTablesName(string connectstring)
        {
            using (OleDbConnection conn = new OleDbConnection(connectstring))
            {
                return GetExcelTablesName(conn);
            }
        }
        public static List<string> GetExcelTablesName(OleDbConnection connection)
        {
            List<string> list = new List<string>();

            if (connection.State == ConnectionState.Closed)
                connection.Open();

            DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    list.Add(ConvertTo<string>(dt.Rows[i][2]));
                }
            }

            return list;
        }
        public static T ConvertTo<T>(object data)
        {
            if (data == null || Convert.IsDBNull(data))
                return default(T);

            object obj = ConvertTo(data, typeof(T));
            if (obj == null)
            {
                return default(T);
            }
            return (T)obj;
        }
        public static object ConvertTo(object data, Type targetType)
        {
            if (data == null || Convert.IsDBNull(data))
            {
                return null;
            }

            Type type2 = data.GetType();
            if (targetType == type2)
            {
                return data;
            }
            if (((targetType == typeof(Guid)) || (targetType == typeof(Guid?))) && (type2 == typeof(string)))
            {
                if (string.IsNullOrEmpty(data.ToString()))
                {
                    return null;
                }
                return new Guid(data.ToString());
            }

            if (targetType.IsEnum)
            {
                try
                {
                    return Enum.Parse(targetType, data.ToString(), true);
                }
                catch
                {
                    return Enum.ToObject(targetType, data);
                }
            }

            if (targetType.IsGenericType)
            {
                targetType = targetType.GetGenericArguments()[0];
            }

            return Convert.ChangeType(data, targetType);
        }
        public static List<string> GetExcelTablesName(string excelPath, ExcelType eType)
        {
            string connectstring = GetExcelConnectstring(excelPath, true, eType);
            return GetExcelTablesName(connectstring);
        }

        public static DataSet ExcelToDataSet(string excelPath, bool header, ExcelType eType)
        {
            string connectstring = GetExcelConnectstring(excelPath, header, eType);
            return ExcelToDataSet(connectstring);
        }
        public static DataSet ExcelToDataSet(string connectstring)
        {
            using (OleDbConnection conn = new OleDbConnection(connectstring))
            {
                DataSet ds = new DataSet();
                List<string> tableNames = GetExcelTablesName(conn);

                foreach (string tableName in tableNames)
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + tableName + "]", conn);
                    adapter.Fill(ds, tableName);
                }
                return ds;
            }
        }
    }
}

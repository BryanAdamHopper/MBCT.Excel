using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//Added Referance: System.Data;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MBCT.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelTools
    {
        /// <summary>
        /// 
        /// </summary>
        public static readonly string Type = "Excel";
        /// <summary>
        /// 
        /// </summary>
        public static readonly List<string> Types = new List<string>
        {
            "xl",
            "xlsx",
            "xlsm",
            "xlsb",
            "xlam",
            "xltx",
            "xltm",
            "xls",
            "xla",
            "xlt",
            "xlm",
            "xlw"
        };
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<string> OpenFileDialogueFilter()
        {
            var types = Types.Aggregate("", (current, type) => current + $"*.{type};");
            types += "|All Files|*.*";
            return new List<string> { $"{Type}", $"{Type} Files|{types}" };
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static OpenFileDialog OpenFileDiag()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            List<string> types = OpenFileDialogueFilter();

            ofd.Title = $"Open {types[0]} File";
            ofd.Filter = $"{types[1]}";
            ofd.FileName = null;

            return ofd;
        }

        /// <summary>
        /// Returns true if file extension matchs common file formats.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IsExcel(string path)
        {
            string ext = Path.GetExtension(path);
            string filesTypes = "*.xl; *.xlsx; *.xlsm; *.xlsb; *.xlam; *.xltx; *.xltm; *.xls; *.xla; *.xlt; *.xlm; *.xlw";

            if (filesTypes.IndexOf(ext) > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        /// <summary>
        /// Returns true if table exists. 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static bool TableExists(string path, string tableName)
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path)))
            {
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                return conn.GetSchema("Tables", new string[4] { null, null, tableName, "TABLE" }).Rows.Count > 0;
            }
        }

        /// <summary>
        /// Returns a list of tables from schema.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static DataTable TableSchema(string path)
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path)))
            {
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                return conn.GetSchema("Tables"); 
            }
        }

        /// <summary>
        /// Creates connection string based on file type and parameters provided by user.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="hasHeader"></param>
        /// <returns></returns>
        public static string GetConnectionString(string path, bool hasHeader = true)
        {
            string PROVIDER = "Microsoft.ACE.OLEDB.12.0";
            string EXTENDED_PROPERTIES = "";
            string HDR = hasHeader ? "YES" : "NO";
            string IMEX = hasHeader ? ";IMEX=1" : "";

            switch (Path.GetExtension(path).ToLower())
            {
                case ".xlsx":
                    EXTENDED_PROPERTIES = "Excel 12.0 Xml";
                    break;
                case ".xlsb":
                    EXTENDED_PROPERTIES = "Excel 12.0";
                    break;
                case ".xlsm":
                    EXTENDED_PROPERTIES = "Excel 12.0 Macro";
                    break;
                case ".xls":
                    EXTENDED_PROPERTIES = "Excel 8.0";
                    break;
                default:
                    EXTENDED_PROPERTIES = "Excel 8.0";
                    break;
            }

            return string.Format($"Provider={PROVIDER};Data Source={path};Extended Properties=\"{EXTENDED_PROPERTIES};HDR={HDR}{IMEX}\";");

            #region
            //string connectionString = "";
            //
            //if (ext.ToLower() == ".xlsx")
            //{
            //    //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;Extended Properties = "Excel 12.0 Xml;HDR=YES;IMEX=1";
            //    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR={1}{2}\";", path, HDR, IMEX);
            //}
            //else if (ext.ToLower() == ".xlsb")
            //{
            //    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR={1}{2}\";", path, HDR, IMEX);
            //}
            //else if (ext.ToLower() == ".xlsm")
            //{
            //    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Macro;HDR={1}{2}\";", path, HDR, IMEX);
            //}
            //else if (ext.ToLower() == ".xls")
            //{
            //    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR={1}{2}\";", path, HDR, IMEX);
            //}
            //else
            //{
            //    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR={1}{2}\";", path, HDR, IMEX);
            //}
            //
            //return connectionString;
            #endregion
        }

        /// <summary>
        /// Reads data from Excel.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataTable Reader(string path, string sheetName)
        {
            string str_BatchExt = Path.GetExtension(path);


            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path)))
            {
                using (OleDbCommand comm = new OleDbCommand(String.Format("Select * from [{0}]", sheetName), conn))
                {
                    using (OleDbDataAdapter da = new OleDbDataAdapter(comm))
                    {
                        //da.SelectCommand = comm;
                        if (conn.State != ConnectionState.Open) { conn.Open(); }

                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        return dt;
                    }
                }
            }


        }

        /// <summary>
        /// Return list of tables in file.
        /// The object here represents the 1st 4 columns, the 4th one being TABLE_TYPE.
        /// TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, TABLE_GUID, DESCRIPTION, TABLE_PROPID, DATE_CREATED, DATE_MODIFIED
        /// </summary>
        /// <param name="path"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DataTable GetTableList(string path, string password = "")
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path)))
            {
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                /*The object here represents the 1st 4 columns, the 4th one being TABLE_TYPE.*/
                /*TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, TABLE_GUID, DESCRIPTION, TABLE_PROPID, DATE_CREATED, DATE_MODIFIED*/
                return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            }
        }
    }
}

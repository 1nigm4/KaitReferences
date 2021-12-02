using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace KaitReference.Services
{
    public class Excel
    {
        public static List<string[]> Export()
        {
            DataTable table = new DataTable();
            using (OleDbConnection conn = new OleDbConnection())
            {
                string filePath = $@"{Environment.CurrentDirectory}\Data\Students.xlsx";
                string fileExtension = Path.GetExtension(filePath);
                conn.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand db = new OleDbCommand())
                {
                    db.CommandText = "Select * from [Реестр контингента$]";

                    db.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = db;
                        da.Fill(table);
                    }
                }
            }

            List<string[]> result = new List<string[]>();
            foreach (DataRow row in table.Rows)
                result.Add(row.ItemArray.Cast<string>().ToArray());
            return result;
        }
    }
}

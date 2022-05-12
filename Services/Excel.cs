using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace KaitReferences.Services
{
    public class Excel
    {
        const string COLUMNS = "`Фамилия`, `Имя`, `Отчество`, `Дата рождения`, `Пол`, `Финансирование (средства обучения)`, `Учебная группа`, `Курс обучения`, `Статус`, `Базовое образование`, `Номер приказа о зачислении`, `Дата приказа о зачислении`, `Дата приема`, `Программа обучения`, `Профессия/специальность`, `Код профессии/специальности`, `Форма обучения`, `Срок обучения`";
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
                    db.CommandText = $"Select {COLUMNS} from [Реестр контингента$]";

                    db.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = db;
                        da.Fill(table);
                    }
                }
            }

            var asd = table.Columns;
            List<string[]> result = new List<string[]>();
            foreach (DataRow row in table.Rows)
                if (row.ItemArray[0] is string)
                    result.Add(row.ItemArray.Cast<string>().ToArray());
            return result;
        }
    }
}

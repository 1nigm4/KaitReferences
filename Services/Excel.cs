using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;

namespace KaitReferences.Services
{
    public class Excel
    {
        const string COLUMNS = "`Фамилия`, `Имя`, `Отчество`, `Дата рождения`, `Пол`, `Финансирование (средства обучения)`, `Учебная группа`, `Курс обучения`, `Статус`, `Базовое образование`, `Номер приказа о зачислении`, `Дата приказа о зачислении`, `Дата приема`, `Программа обучения`, `Профессия/специальность`, `Код профессии/специальности`, `Форма обучения`, `Срок обучения`";
        
        static Excel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        
        public static List<string[]> Export()
        {
            string excelPath = Path.Combine(Environment.CurrentDirectory, @"Data\Students.xlsx");
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var sheet = package.Workbook.Worksheets.First();
                    var rows = sheet.Cells.Where(cell => cell != null)
                        .GroupBy(cell => cell.EntireRow.StartRow)
                        .Where(row => row.FirstOrDefault().Value != null)
                        .Select(row => row.Where(cell => COLUMNS.Contains(cell.EntireColumn.Range.Text))
                            .DistinctBy(cell => cell.EntireColumn.Range.Text)
                            .Select(cell => cell.Value.ToString())
                            .ToArray())
                        .Skip(1)
                        .ToList();

                    return rows;
                }
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show("Необходимо загрузить базу АИС", "Экспорт данных");
                return null;
            }
        }
    }
}

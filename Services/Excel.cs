using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using MsExcel = Microsoft.Office.Interop.Excel;
namespace KaitReference.Services
{
    static class Excel
    {
        public static MsExcel.Application App;

        static Excel()
        {
            App = new MsExcel.Application();
        }
        public static IEnumerable<string[]> Export()
        {
            App.Workbooks.Open(@$"{Directory.GetCurrentDirectory()}\Data\Students.xlsx");
            MsExcel.Worksheet sheet = (MsExcel.Worksheet)App.Worksheets.get_Item(1);
            int columns = sheet.UsedRange.Columns.Count;
            int rows = sheet.UsedRange.Rows.Count;

            for (int i = 2; i <= rows; i++)
            {
                string data = string.Empty;
                for (int j = 1; j <= columns; j++)
                {
                    data += ((MsExcel.Range)sheet.Cells[i, j]).Value.ToString() + ";";
                }
                yield return data.Split(';');
            }
        }
    }
}

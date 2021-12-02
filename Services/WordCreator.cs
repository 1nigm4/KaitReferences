using KaitReference.Models;
using KaitReference.Views.Windows;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Windows;

namespace KaitReference.Services
{
    class WordCreator
    {
        public static void CreateReference(Person person)
        {
            Word.Application word = new Word.Application();
            word.Visible = (bool)MainWindow.WordVisible.IsChecked;
            Word.Document doc = word.Documents.Add($@"{Directory.GetCurrentDirectory()}\Data\Templates\Reference.docx");
            doc.Bookmarks["ФИО"].Range.Text = person.ToString();
            doc.Bookmarks["ДатаРождения"].Range.Text = person.BirthDate.ToShortDateString();
            doc.Bookmarks["Курс"].Range.Text = person.Education.Course.ToString();
            doc.Bookmarks["ФормаОбучения"].Range.Text = person.Education.Form.Contains("Очная") ? "очной" : person.Education.Form.Contains("Заочная") ? "заочной" : "очно-заочной";
            doc.Bookmarks["НомерПриказа"].Range.Text = person.Education.OrderNumber;
            doc.Bookmarks["ДатаПриказа"].Range.Text = person.Education.OrderDate.ToShortDateString();
            doc.Bookmarks["ДатаПриема"].Range.Text = person.Education.AdmissionDate.ToShortDateString();
            doc.Bookmarks["КодСпециальности"].Range.Text = person.Education.SpecialityCode;
            doc.Bookmarks["ПрограммаОбучения"].Range.Text = person.Education.Program;
            doc.Bookmarks["Специальность"].Range.Text = person.Education.Speciality;
            doc.Bookmarks["Финансирование"].Range.Text = person.Education.Financing == "Бюджет" ? "бюджетных ассигнований" : "средств юридических лиц";
            doc.Bookmarks["ДатаОкончания"].Range.Text = person.Education.EndDate.ToShortDateString();
            doc.Bookmarks["Назначение"].Range.Text = person.Reference.Assignment;
            doc.Bookmarks["Площадка"].Range.Text = person.Education.Area;
            doc.Bookmarks["ДатаВыдачи"].Range.Text = DateTime.Now.ToShortDateString();
            doc.Bookmarks["НомерВыдачи"].Range.Text = GoogleSheets.GetLastReferenceIndex();

            if (!word.Visible)
            {
                string saveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Готовые справки";
                if (!Directory.Exists(saveFilePath))
                    Directory.CreateDirectory(saveFilePath);
                doc.SaveAs2($"{saveFilePath}\\{person} {DateTime.Now.ToShortDateString()} - {GoogleSheets.GetLastReferenceIndex()}.docx", Word.WdSaveFormat.wdFormatDocumentDefault);
                word.Quit();
                MessageBox.Show($"Файл успешно сохранен в {saveFilePath}", "Информация",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                GoogleSheets.AddReference(person);
            }   
        }

        public static void CreateRectal(Person person)
        {
            Word.Application word = new Word.Application();
            word.Visible = (bool)MainWindow.WordVisible.IsChecked;
            Word.Document doc = word.Documents.Add($@"{Directory.GetCurrentDirectory()}\Data\Templates\Rectal.docx");
            doc.Bookmarks["ФИО"].Range.Text = person.ToString();
            doc.Bookmarks["ДатаРождения"].Range.Text = $"{person.BirthDate.Year}";
            doc.Bookmarks["ДатаПриема"].Range.Text = $"{person.Education.AdmissionDate.Year}";
            doc.Bookmarks["УровеньОбразования"].Range.Text = person.Education.Base;
            doc.Bookmarks["НомерПриказа"].Range.Text = person.Education.OrderNumber;
            doc.Bookmarks["Курс"].Range.Text = person.Education.Course.ToString();
            doc.Bookmarks["БазовыйКодСпециальности"].Range.Text = person.Education.BaseSpecialityCode;
            doc.Bookmarks["БазоваяСпециальность"].Range.Text = person.Education.BaseSpeciality;
            doc.Bookmarks["ФормаОбучения"].Range.Text = person.Education.Form == "Очная" ? "очной" : person.Education.Form == "Заочная" ? "заочной" : "очно-заочной";
            doc.Bookmarks["КодСпециальности"].Range.Text = person.Education.SpecialityCode;
            doc.Bookmarks["Специальность"].Range.Text = person.Education.Speciality;
            doc.Bookmarks["ДатаОкончания"].Range.Text = $"{person.Education.EndDate.Year}";
            if (!word.Visible)
                word.Quit();
        }
    }
}

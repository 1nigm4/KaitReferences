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
        public static void CreateReference(Person user)
        {
            Word.Application word = new Word.Application();
            word.Visible = (bool)MainWindow.WordVisible.IsChecked;
            Word.Document doc = word.Documents.Add($@"{Directory.GetCurrentDirectory()}\Data\Templates\Reference.docx");
            doc.Bookmarks["ФИО"].Range.Text = user.ToString();
            doc.Bookmarks["ДатаРождения"].Range.Text = user.BirthDate.ToShortDateString();
            doc.Bookmarks["Курс"].Range.Text = user.Education.Course.ToString();
            doc.Bookmarks["ФормаОбучения"].Range.Text = user.Education.Form == "Очная" ? "очной" : user.Education.Form == "Заочная" ? "заочной" : "очно-заочной";
            doc.Bookmarks["НомерПриказа"].Range.Text = user.Education.OrderNumber;
            doc.Bookmarks["ДатаПриказа"].Range.Text = user.Education.OrderDate.ToShortDateString();
            doc.Bookmarks["ДатаПриема"].Range.Text = user.Education.AdmissionDate.ToShortDateString();
            doc.Bookmarks["КодСпециальности"].Range.Text = user.Education.SpecialityCode;
            doc.Bookmarks["ПрограммаОбучения"].Range.Text = user.Education.Program;
            doc.Bookmarks["Специальность"].Range.Text = user.Education.Speciality;
            doc.Bookmarks["Финансирование"].Range.Text = user.Education.Financing == "Бюджет" ? "бюджетных ассигнований" : "средств юридических лиц";
            doc.Bookmarks["ДатаОкончания"].Range.Text = user.Education.EndDate.ToShortDateString();
            doc.Bookmarks["Назначение"].Range.Text = user.Reference.Assignment;
            doc.Bookmarks["Площадка"].Range.Text = user.Education.Area;

            if (!word.Visible)
            {
                string saveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Готовые справки";
                if (!Directory.Exists(saveFilePath))
                    Directory.CreateDirectory(saveFilePath);
                doc.SaveAs2($"{saveFilePath}\\{user} {DateTime.Now.ToShortDateString()}.docx", Word.WdSaveFormat.wdFormatDocumentDefault);
                word.Quit();
                MessageBox.Show($"Файл успешно сохранен в {saveFilePath}", "Информация",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }   
        }

        public static void CreateRectal(Person user)
        {
            Word.Application word = new Word.Application();
            word.Visible = (bool)MainWindow.WordVisible.IsChecked;
            Word.Document doc = word.Documents.Add($@"{Directory.GetCurrentDirectory()}\Data\Templates\Rectal.docx");
            doc.Bookmarks["ФИО"].Range.Text = user.ToString();
            doc.Bookmarks["ДатаРождения"].Range.Text = $"{user.BirthDate.Year}";
            doc.Bookmarks["Пол"].Range.Text = user.Gender == "Мужской" ? "он" : "она";
            doc.Bookmarks["ДатаПриема"].Range.Text = $"{user.Education.AdmissionDate.Year}";
            doc.Bookmarks["УровеньОбразования"].Range.Text = user.Education.Base;
            doc.Bookmarks["НомерПриказа"].Range.Text = user.Education.OrderNumber;
            doc.Bookmarks["ДатаПриказа"].Range.Text = user.Education.OrderDate.ToLongDateString();
            doc.Bookmarks["Курс"].Range.Text = user.Education.Course.ToString();
            doc.Bookmarks["ФормаОбучения"].Range.Text = user.Education.Form == "Очная" ? "очной" : user.Education.Form == "Заочная" ? "заочной" : "очно-заочной";
            doc.Bookmarks["ПрограммаОбучения"].Range.Text = user.Education.Program;
            doc.Bookmarks["КодСпециальности"].Range.Text = user.Education.SpecialityCode;
            doc.Bookmarks["Специальность"].Range.Text = user.Education.Speciality;
            doc.Bookmarks["ДатаОкончания"].Range.Text = $"{user.Education.EndDate.Year}";
            if (!word.Visible)
                word.Quit();
        }
    }
}

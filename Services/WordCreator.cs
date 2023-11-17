using KaitReferences.Extensions;
using KaitReferences.Models;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace KaitReferences.Services
{
    class WordCreator
    {
        static string _DIRECTORY = Directory.GetCurrentDirectory();

        public static void CreateReference(Person person)
        {
            person.Reference.ReferenceType = ReferenceType.Reference;
            string program = person.Education.Program == "Профессия" ? "квалифицированных рабочих, служащих" : "специалистов среднего звена";
            string number = person.Education.OrderNumber.Split("/")[0] + "/лу";

            using (var document = DocX.Load(@$"{_DIRECTORY}\Data\Templates\Reference.docx"))
            {
                document.WithUnderline(false);
                document.SetText("ФИО", person.FIO);
                document.SetText("ДатаРождения", $"{person.BirthDate:d}");
                document.SetText("Курс", $"{person.Education.Course}");
                document.SetText("ФормаОбучения", person.Education.Form);
                document.SetText("Подготовка", program);
                document.SetText("НомерПриказа", number);
                document.SetText("ДатаПриказа", $"{person.Education.OrderDate:d}");
                document.SetText("ДатаПриема", $"{person.Education.AdmissionDate:d}");
                document.SetText("КодСпециальности", person.Education.SpecialityCode);
                document.SetText("ПрограммаОбучения", person.Education.Program);
                document.SetText("Специальность", person.Education.Speciality);
                document.SetText("Финансирование", person.Education.Financing);
                document.SetText("ПериодОбучения", person.Education.Period);
                document.SetText("ДатаОкончания", $"{person.Education.EndDate:d}");
                document.SetText("Назначение", person.Reference.Assignment);
                document.SetText("ДатаВыдачи", $"{DateTime.Now:d}");
                document.SetText("НомерВыдачи", GoogleSheets.GetLastReferenceIndex(person));
                document.SetText("Исполнитель", GoogleSheets.Executor);
                document.SetText("Площадка", person.Education.Area);
                SaveFile(document, person);
            }
        }

        public static void CreateRectal(Person person)
        {
            person.Reference.ReferenceType = ReferenceType.Rectal;
            string program = person.Education.Program == "Профессия" ? "профессии" : "специальности";

            using (var document = DocX.Load(@$"{_DIRECTORY}\Data\Templates\Rectal.docx"))
            {
                document.WithUnderline(true);
                document.SetText("ФИО", person.FIO);
                document.SetText("ДатаРождения", $"{person.BirthDate.Year}");
                document.SetText("ДатаПриема", $"{person.Education.AdmissionDate.Year}");
                document.SetText("УровеньОбразования", person.Education.Base);
                document.SetText("Курс", $"{person.Education.Course}");
                document.SetText("БазовыйКодСпециальности", person.Education.BaseSpecialityCode);
                document.SetText("БазоваяСпециальность", person.Education.BaseSpeciality);
                document.SetText("ФормаОбучения", person.Education.Form);
                document.SetText("ПрограммаОбучения", program);
                document.SetText("ФормаОбучения1", person.Education.Form);
                document.SetText("КодСпециальности", person.Education.SpecialityCode);
                document.SetText("Специальность", person.Education.Speciality);
                document.SetText("ДатаОкончания", $"{person.Education.EndDate.Year}");
                document.SetText("ПериодОбучения", person.Education.Period);
                document.SetText("ДатаВыдачи", $"{DateTime.Now:d}");
                document.SetText("НомерВыдачи", GoogleSheets.GetLastReferenceIndex(person));
                document.SetText("Исполнитель", GoogleSheets.Executor);
                document.SetText("Площадка", person.Education.Area);
                SaveFile(document, person);
            }
        }

        private static void SaveFile(DocX document, Person person)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Готовые справки";
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string filePath = $"{folderPath}\\{person} {DateTime.Now.ToShortDateString()} - {GoogleSheets.GetLastReferenceIndex(person)}.docx";
            document.SaveAs(filePath);
            MessageBoxResult result = MessageBox.Show($"Справка успешно сохранена по пути {folderPath}. Открыть справку?", "Информация",
                MessageBoxButton.YesNo, MessageBoxImage.Information);

            if (result == MessageBoxResult.Yes)
            {
                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
        }
    }
}

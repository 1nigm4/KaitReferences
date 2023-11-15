using KaitReferences.Commands;
using KaitReferences.Models;
using KaitReferences.Services;
using KaitReferences.ViewModels.Base;
using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace KaitReferences.ViewModels
{
    class MainWindowViewModel : ViewModel
    {
        private List<Person> persons;
        public List<Person> Persons
        {
            get => persons;
            set => Set(ref persons, value);
        }
        private Person selectedPerson;
        public Person SelectedPerson
        {
            get => selectedPerson;
            set => Set(ref selectedPerson, value);
        }

        private string errorReason;
        public string ErrorReason
        {
            get => errorReason;
            set => Set(ref errorReason, value);
        }

        #region Commands
        public ICommand CreateReferenceCommand { get; }
        private void OnCreateReferenceCommandExecuted(object p) => WordCreator.CreateReference(SelectedPerson);
        private bool CanCreateReferenceCommandExecute(object p)
        {
            if (SelectedPerson?.Education?.Status is null or "В академическом отпуске")
            {
                ErrorReason = SelectedPerson?.Education?.Status ?? "Нет в базе";
                return false;
            }

            ErrorReason = string.Empty;
            return true;
        }
        public ICommand CreateRectalCommand { get; }
        private void OnCreateRectalCommandExecuted(object p) => WordCreator.CreateRectal(SelectedPerson);
        private bool CanCreateRectalCommandExecute(object p)
        {
            if (SelectedPerson == null || SelectedPerson.Gender is not ("Мужской" or "м.")) return false;
            if (SelectedPerson.Education.Form == "заочной")
            {
                ErrorReason = "Заочник";
                return false;
            }
            return CanCreateReferenceCommandExecute(null);
        }
        public ICommand SaveReferenceStatusCommand { get; }
        private void OnSaveReferenceStatusCommandExecuted(object p)
        {
            int index = Persons.IndexOf(SelectedPerson);
            GoogleSheets.SaveStatusChanges(index, SelectedPerson.Reference.Status);
            if (SelectedPerson.Reference.Status.Contains("да"))
                GoogleSheets.AddReference(SelectedPerson);
        }
        private bool CanSaveReferenceStatusCommandExecute(object p) => SelectedPerson != null && !string.IsNullOrWhiteSpace(SelectedPerson.Reference.Status);
        public ICommand SynchronizationCommand { get; }
        private void OnSynchronizationCommandExecuted(object p) => Synchronization();
        private bool CanSynchronizationCommandExecute(object p) => true;
        public ICommand UploadStudentsCommand { get; }
        private void OnUploadStudentsCommandExecuted(object p)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files 2007+|*.xlsx" + "|" + "Excel Files 93-2000|*.xls";
            if (fileDialog.ShowDialog() == true)
            {
                FileInfo data = new FileInfo(fileDialog.FileName);
                if (data.Extension == ".xls")
                {
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(data.FullName);
                    workbook.SaveToFile("Students.xlsx", ExcelVersion.Version2016);
                    data = new FileInfo(workbook.FileName);
                }
                data.CopyTo(Environment.CurrentDirectory + "\\Data\\Students.xlsx", true);
                Synchronization();
            }
        }
        private bool CanUploadStudentsCommandExecute(object p) => true;
        #endregion

        public MainWindowViewModel()
        {
        connect: // Trying to connect to Google Sheets
            var isConnected = GoogleSheets.Connect();
            if (!isConnected)
            {
                var result = MessageBox.Show("Ошибка синхронизации с Google Sheets. Повторить авторизацию?", "Google Sheets", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                    goto connect;
                Environment.Exit(0);
            }

            CreateReferenceCommand = new LambdaCommand(OnCreateReferenceCommandExecuted, CanCreateReferenceCommandExecute);
            CreateRectalCommand = new LambdaCommand(OnCreateRectalCommandExecuted, CanCreateRectalCommandExecute);
            SaveReferenceStatusCommand = new LambdaCommand(OnSaveReferenceStatusCommandExecuted, CanSaveReferenceStatusCommandExecute);
            SynchronizationCommand = new LambdaCommand(OnSynchronizationCommandExecuted, CanSynchronizationCommandExecute);
            UploadStudentsCommand = new LambdaCommand(OnUploadStudentsCommandExecuted, CanUploadStudentsCommandExecute);

            SynchronizationCommand.Execute(null);
        }

        private async void Synchronization()
        {
            SelectedPerson = null;
            Persons = GoogleSheets.ExportReferences();
            Persons.Reverse();
            await Task.Run(() => GetMoreInformation());
        }

        private void GetMoreInformation()
        {
            List<string[]> table = Excel.Export();
            if (table == null) return;

            Parallel.ForEach(Persons, new ParallelOptions { MaxDegreeOfParallelism = -1 }, person =>
            {
                string[] data = table.Find(d => PersonWithGroup(person, d));
                if (data == null) return;

                string[] fio = GetFIO(data[0]);
                person.LastName = fio[0];
                person.Name = fio[1];
                person.Patronymic = fio[2];
                person.BirthDate = DateTime.Parse(data[1]);
                person.Gender = data[2];
                person.Education.Status = data[3];
                person.Education.Financing = data[4] == "Бюджет" ? "бюджетных ассигнований" : "средств физических лиц";
                person.Education.Group = data[5].Split('.')[0];
                person.Education.Area = GetArea(data[5]);
                int course = int.Parse(data[6].Split()[0]);
                person.Education.Course = course;
                person.Education.OrderNumber = data[7];
                person.Education.OrderDate = DateTime.Parse(data[8]);
                person.Education.AdmissionDate = DateTime.Parse(data[9]);
                person.Education.Program = data[10].Contains("ППССЗ") ? "Специальность" : "Профессия";
                person.Education.Speciality = data[11];
                person.Education.SpecialityCode = data[12];
                string[] baseSpeciality = GoogleSheets.GetBaseSpecialityCode(data[12]);
                person.Education.BaseSpeciality = baseSpeciality[1];
                person.Education.BaseSpecialityCode = baseSpeciality[0];
                person.Education.Form = data[13].Contains("Очная") ? "очной" : data[16].Contains("Заочная") ? "заочной" : "очно-заочной";
                string period = data[14];
                person.Education.Period = period;
                person.Education.EndDate = GetEndDate(period, course);
                person.Education.Base = data[15];
            });
        }

        private string GetArea(string area)
        {
            return Regex.Matches(area, @"\d")[2].Value.Last() switch
            {
                '1' => "датахаб",
                '2' => "диджитал",
                '3' => "авто",
                '4' => "техно",
                '5' => "кибер",
                '6' => "моссовет"
            };
        }

        private DateTime GetEndDate(string periodDate, int course)
        {
            int period = periodDate switch
            {
                "10м" => 1,
                "3г 10м" => 4,
                "4г 10м" => 5,
                _ => 3 // 2г 4м и 2г 10м
            };

            DateTime date = DateTime.Now;
            int halfYear = date.Month < 9 ? 0 : 1;
            int endDateYear = date.Year + (period - course) + halfYear;
            return DateTime.Parse($"30.06.{endDateYear}");
        }

        private bool PersonWithFIO(Person person, string[] aisPerson)
        {
            if (person.FIO == "Монахов Андрей Антонович" || aisPerson[0] == "Монахов Андрей Антонович")
                _ = 1;
            return aisPerson[0] == person.FIO;
        }
            

        private bool PersonWithGroup(Person person, string[] aisPerson) =>
            PersonWithFIO(person, aisPerson) &&
            aisPerson[5].ToLower().Contains(person.Education.Group.ToLower());

        private string[] GetFIO(string FIO)
        {
            string[] fio = FIO.Split();

            string lastName = fio[0];
            string firstName = fio[1];
            string patronymic = string.Join(" ", fio.Skip(2))
                .Replace("-", string.Empty);

            return new string[] { lastName, firstName, patronymic };
        }
    }
}

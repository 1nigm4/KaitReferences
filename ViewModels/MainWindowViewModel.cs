using KaitReference.Commands;
using KaitReference.Models;
using KaitReference.Services;
using KaitReference.ViewModels.Base;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace KaitReference.ViewModels
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

        #region Commands
        public ICommand CreateReferenceCommand { get; }
        private void OnCreateReferenceCommandExecuted(object p) => WordCreator.CreateReference(SelectedPerson);
        private bool CanCreateReferenceCommandExecute(object p) => SelectedPerson != null;
        public ICommand CreateRectalCommand { get; }
        private void OnCreateRectalCommandExecuted(object p) => WordCreator.CreateRectal(SelectedPerson);
        private bool CanCreateRectalCommandExecute(object p) => SelectedPerson != null && SelectedPerson.Gender == "Мужской";
        public ICommand SaveReferenceStatusCommand { get; }
        private void OnSaveReferenceStatusCommandExecuted(object p)
        {
            int index = Persons.IndexOf(SelectedPerson);
            GoogleSheets.SaveStatusChanges(index, SelectedPerson.Reference.Status);
        }
        private bool CanSaveReferenceStatusCommandExecute(object p) => SelectedPerson != null && !string.IsNullOrWhiteSpace(SelectedPerson.Reference.Status);
        public ICommand SynchronizationCommand { get; }
        private void OnSynchronizationCommandExecuted(object p) => Synchronization();
        private bool CanSynchronizationCommandExecute(object p) => true;
        public ICommand UploadStudentsCommand { get; }
        private void OnUploadStudentsCommandExecuted(object p)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files | *.xlsx";
            if (fileDialog.ShowDialog() == true)
            {
                FileInfo data = new FileInfo(fileDialog.FileName);
                data.CopyTo(Environment.CurrentDirectory + "\\Data\\Students.xlsx", true);
            }
        }
        private bool CanUploadStudentsCommandExecute(object p) => true;
        #endregion

        public MainWindowViewModel()
        {
            connect: // Попытка переподключения к Google Sheets
            var isConnected = GoogleSheets.Connect().Result;
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
            await Task.Run(() => GetMoreInformation());
        }

        private void GetMoreInformation()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Parallel.ForEach(Excel.Export(), data =>
            {
                Person person = Persons.Find(p => p.LastName == data[0] &
                                               p.Name == data[1] &
                                               p.Patronymic == data[2]);

                if (person != null)
                {
                    person.BirthDate = DateTime.Parse(data[3]);
                    person.Gender = data[4];
                    person.Education.Financing = data[5];
                    person.Education.Group = data[6].Split('.')[0];
                    person.Education.Area = Regex.Matches(data[6], @"\d")[2].Value.Last() switch
                    {
                        '1' => "юниор",
                        '2' => "1М",
                        '3' => "авто",
                        '4' => "техно",
                        '5' => "бтм",
                        '6' => "моссовет",
                        _ => ""
                    };
                    person.Education.Course = data[7] == "I" ? 1 : data[7] == "II" ? 2 : data[7] == "III" ? 3 : 4;
                    person.Education.Status = data[8];
                    person.Education.OrderNumber = data[9];
                    person.Education.OrderDate = DateTime.Parse(data[10]);
                    person.Education.AdmissionDate = DateTime.Parse(data[11]);
                    person.Education.Program = data[12].Contains("ППССЗ") ? "Специальность" : "Профессия";
                    person.Education.Speciality = data[13];
                    person.Education.SpecialityCode = data[14];
                    person.Education.Form = data[15];
                    person.Education.Period = data[16].Contains("2г") ? 3 : 4;
                    person.Education.Base = data[17].Split('(')[0]; // Убираем пояснение (5-9 класс)
                    int endDateYear = person.Education.AdmissionDate.Year + person.Education.Period;
                    person.Education.EndDate = DateTime.Parse($"30.06.{endDateYear}");
                }
            });
            stopwatch.Stop();
            MessageBox.Show(stopwatch.Elapsed.TotalSeconds.ToString());
        }
    }
}

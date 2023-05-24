using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.PeopleService.v1;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using KaitReferences.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace KaitReferences.Services
{
    class GoogleSheets
    {
        public static string Executor;
        private static SheetsService sheetsService;
        private static readonly string appName;
        private static readonly string clientId;
        private static readonly string clientSecret;
        private static string sheetName;
        private static string spreadSheetId;
        private static int? referenceJournalId = 1618961239;
        private static int? rectalJournalId = 102128820;
        private static int? directionsId = 1755638865;
        private static List<IList<object>> specialityCodes;

        static GoogleSheets()
        {
            appName = Properties.Resources.AppName;
            clientId = Properties.Resources.ClientId;
            clientSecret = Properties.Resources.ClientSecret;
            sheetName = Properties.Resources.SheetName;
        }

        public static bool Connect()
        {
            string[] scopes = new string[] { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets, PeopleServiceService.Scope.UserinfoProfile };
            UserCredential credentials = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            }, scopes,
            Environment.UserName, CancellationToken.None, new FileDataStore("MyAppsToken")).Result;

            DriveService driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credentials,
                ApplicationName = appName
            });

            sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credentials,
                ApplicationName = appName
            });

            PeopleServiceService peopleService = new PeopleServiceService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credentials
            });

            var peopleRequest = peopleService.People.Get("people/me");
            peopleRequest.PersonFields = "names";
            var person = peopleRequest.Execute();
            string familyName = person.Names[0].FamilyName;
            string[] givenName = person.Names[0].GivenName.Split();
            Executor = $"{familyName} {givenName[0][0]}.{givenName[1][0]}.";

            var driveRequest = driveService.Files.List();
            var response = driveRequest.Execute();

            spreadSheetId = response.Files.FirstOrDefault(f => f.Name == sheetName)?.Id;

            if (string.IsNullOrEmpty(spreadSheetId)) return false;

            var sheetRequest = sheetsService.Spreadsheets.Get(spreadSheetId);
            var sheetResponse = sheetRequest.Execute();

            sheetName = sheetResponse.Sheets[0].Properties.Title;

            return true;
        }

        public static List<Person> ExportReferences()
        {
            var request = sheetsService.Spreadsheets.Values.Get(spreadSheetId, sheetName);
            var sheet = request.Execute().Values;
            List<Person> references = new List<Person>();
            for (int n = 1; n < sheet.Count; n++)
            {
                var reference = sheet[n];
                while (reference.Count < 17) reference.Add(string.Empty); // Google sheets skiping last cells without info
                Person person = new Person()
                {
                    EmailAddress = (string)reference[1],
                    LastName = (string)reference[3],
                    Name = (string)reference[4],
                    Patronymic = (string)reference[5],
                    Phone = (string)reference[12],
                    Email = (string)reference[13],
                    Education = new Education()
                    {
                        Area = (string)reference[2],
                        Group = (string)reference[6]
                    },
                    Reference = new Reference()
                    {
                        Date = DateTime.Parse((string)reference[0]),
                        Type = ((string)reference[7]).Split()[0],
                        Count = int.Parse((string)reference[8]),
                        Assignment = (string)reference[9],
                        Period = (string)reference[10],
                        Form = (string)reference[11],
                        Note = (string)reference[14],
                        Status = (string)reference[16]
                    }
                };
                references.Add(person);
            }
            return references;
        }

        private const string STATUSCOLUMN = "Q"; // Column with status
        private const int INDENT = 2; // Skip 2 first lines of Google sheets
        public static void SaveStatusChanges(int index, string value)
        {
            var range = sheetName + "!" + STATUSCOLUMN + (index + 1 + INDENT);
            ValueRange data = new ValueRange();
            data.Values = new List<IList<object>>() { new List<object>() { value } };

            var request = sheetsService.Spreadsheets.Values.Update(data, spreadSheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }
        public static void AddReference(Person person)
        {
            string sheetName = GetSheetName(person.Reference);
            var request = sheetsService.Spreadsheets.Values.Get(spreadSheetId, sheetName);
            int index = request.Execute().Values.Count + 1;

            var range = sheetName + $"!A{index}:F{index}"; // 'A' column with number of reference; 'F' is last column with note
            ValueRange data = new ValueRange
            {
                Values = new List<IList<object>>()
                {
                    new List<object>()
                    {
                        GetLastReferenceIndex(person),
                        DateTime.Now.ToShortDateString(),
                        person.LastName,
                        person.Name,
                        person.Patronymic,
                        person.Reference.Assignment == "В военный комиссариат" ? "В военный комиссариат" : "Об обучении"
                    }
                }
            };
            var appendRequest = sheetsService.Spreadsheets.Values.Append(data, spreadSheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
        }

        public static string GetLastReferenceIndex(Person person)
        {
            string sheetName = GetSheetName(person.Reference);
            var request = sheetsService.Spreadsheets.Values.Get(spreadSheetId, sheetName);
            var sheet = request.Execute().Values;
            var indexReference = sheet.Last()[0].ToString().All(char.IsNumber) ? sheet.Last()[0] : default(int);
            return (Convert.ToInt32(indexReference) + 1).ToString();
        }

        private static string GetSheetName(Reference reference)
        {
            int? sheetId = reference.ReferenceType == ReferenceType.Rectal || reference.Assignment == "В военный комиссариат" ? rectalJournalId : referenceJournalId;
            string sheetName = sheetsService.Spreadsheets
                .Get(spreadSheetId)
                .Execute()
                .Sheets.First(s => s.Properties.SheetId == sheetId || s.Properties.Title == (sheetId == rectalJournalId ? "Журнал ВК" : "Журнал справок")).Properties.Title;
            return sheetName;
        }

        public static string[] GetBaseSpecialityCode(string specialityCode)
        {
            if (specialityCodes == null)
            {
                string sheetName = sheetsService.Spreadsheets
                    .Get(spreadSheetId)
                    .Execute().Sheets
                    .First(s => s.Properties.SheetId == directionsId || s.Properties.Title == "Направления").Properties.Title;
                var request = sheetsService.Spreadsheets.Values.Get(spreadSheetId, sheetName);
                specialityCodes = new List<IList<object>>(request.Execute().Values);
            }
            
            string[] result = new string[2];
            for (int i = 1; i < specialityCodes.Count; i++)
            {
                if (specialityCode.Split('.')[0] == specialityCodes[i][1].ToString().Split('.')[0])
                {
                    result[0] = specialityCodes[i][1].ToString();
                    result[1] = specialityCodes[i][2].ToString();
                    break;
                }
            }
            return result;
        }
    }
}
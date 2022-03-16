using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
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
        private static SheetsService sheetsService;
        private static readonly string appName;
        private static readonly string clientId;
        private static readonly string clientSecret;
        private static string sheetName;
        private static string sheetId;

        static GoogleSheets()
        {
            appName = Properties.Resources.AppName;
            clientId = Properties.Resources.ClientId;
            clientSecret = Properties.Resources.ClientSecret;
            sheetName = Properties.Resources.SheetName;
        }

        public static bool Connect()
        {
            string[] scopes = new string[] { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
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

            var request = driveService.Files.List();
            var response = request.Execute();

            sheetId = response.Files.FirstOrDefault(f => f.Name == sheetName)?.Id;

            if (string.IsNullOrEmpty(sheetId)) return false;

            var sheetRequest = sheetsService.Spreadsheets.Get(sheetId);
            var sheetResponse = sheetRequest.Execute();

            sheetName = sheetResponse.Sheets[0].Properties.Title;

            return true;
        }

        public static List<Person> ExportReferences()
        {
            var request = sheetsService.Spreadsheets.Values.Get(sheetId, sheetName);
            var sheet = request.Execute().Values;
            List<Person> references = new List<Person>();
            for (int n = 2; n < sheet.Count; n++)
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
                        Date = DateTime.Parse((string) reference[0]),
                        Type = ((string) reference[7]).Split()[0],
                        Count = int.Parse((string) reference[8]),
                        Assignment = (string) reference[9],
                        Period = (string) reference[10],
                        Form = (string) reference[11],
                        Note = (string) reference[14],
                        Status = (string) reference[16]
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

            var request = sheetsService.Spreadsheets.Values.Update(data, sheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }
        public static void AddReference(Person person)
        {
            int sheetIndex = person.Reference.Assignment == "В военный комиссариат" ? 2 : 1;
            string sheetName = sheetsService.Spreadsheets.Get(sheetId).Execute().Sheets[sheetIndex].Properties.Title;

            var request = sheetsService.Spreadsheets.Values.Get(sheetId, sheetName);
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
            var appendRequest = sheetsService.Spreadsheets.Values.Append(data, sheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
        }

        public static string GetLastReferenceIndex(Person person)
        {
            int sheetIndex = person.Reference.Assignment == "В военный комиссариат" ? 2 : 1;
            string sheetName = sheetsService.Spreadsheets.Get(sheetId).Execute().Sheets[sheetIndex].Properties.Title;
            var request = sheetsService.Spreadsheets.Values.Get(sheetId, sheetName);
            var sheet = request.Execute().Values;
            var indexReference = sheet.Last()[0].ToString().All(char.IsNumber) ? sheet.Last()[0] : default(int);
            return (Convert.ToInt32(indexReference) + 1).ToString();
        }


        private static List<IList<object>> specialityCodes;
        public static string[] GetBaseSpecialityCode(string specialityCode)
        {
            if (specialityCodes == null)
            {
                string sheetName = sheetsService.Spreadsheets.Get(sheetId).Execute().Sheets[2].Properties.Title;
                var request = sheetsService.Spreadsheets.Values.Get(sheetId, sheetName);
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
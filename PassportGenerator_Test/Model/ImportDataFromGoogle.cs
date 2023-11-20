using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using GoogleSheetsDownloader.Properties;
using GoogleSheetsDownloader.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;


namespace GoogleSheetsDownloader.Model {
    internal class ImportDataFromGoogle {

        /// <summary>
        /// Конструктор класса GoogleSettings в котором хранятся данные для связи с Google Sheets
        /// </summary>
        /// <param name="spreadsheetId">ID таблицы из которой берутся данные</param>
        /// <param name="fileName">Имя файла json</param>
        /// <param name="ApplicationName">Имя проекта, созданного в Google Cloud</param>
        public ImportDataFromGoogle(string spreadsheetId, string fileName, string ApplicationName) {
            this.fileName = fileName;
            this.spreadsheetId = spreadsheetId;
            this.ApplicationName = ApplicationName;
            
        }

        // Определяем права доступа пользователя к таблице
        internal string[] Scopes { get; set; } = { SheetsService.Scope.SpreadsheetsReadonly };

        // Указываем название проекта, который создали в Google Cloud
        internal string ApplicationName { get; set; }

        // Указываем имя json файла
        public string fileName { get; set; } // "passportgenerator-403707-030caf5e945d.json";

        // Указываем ID таблицы из которой берем данные
        internal string spreadsheetId { get; set; }

        internal string json_path(string fileName) => Path.GetFullPath(fileName);


        /// <summary>
        /// Запуск метода загрузки данных
        /// </summary>
        /// <param name="range"></param>
        /// <param name="excelfile_name"></param>
        internal void ImportFromGoogle(string range, string excelfile_name) {
            IList<IList<Object>> values = GetGoogleSheetsValue(json_path, spreadsheetId, range);
            ExcelWorkFunc.FillInAnExcel(values, excelfile_name);
        }

        /// <summary>
        /// определение ID таблицы из которой берутся данные. 
        /// </summary>
        /// <param name="range">Строка с диапазоном требуемых значений</param>
        /// <param name=""></param>
        /// <returns></returns>
        internal IList<IList<Object>> GetGoogleSheetsValue(string json_path, string spreadsheetId, string range) {

            var service = CreateConnectionGoogle();
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Выполнение запроса и получение данных из Google Sheets
            IList<IList<Object>> values = request.Execute().Values;

            return values;
        }

        /// <summary>
        /// Получаем список листов из Google
        /// </summary>
        /// <returns></returns>
        internal List<string> GetListsFromSheets() {
            var service = CreateConnectionGoogle();

            // Получение информации о таблице
            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();

            // Получение списка листов
            var sheetTitles = spreadsheet.Sheets.Select(s => s.Properties.Title).ToList();
            return sheetTitles;
        }


        /// <summary>
        /// Установка связи с таблицей Google
        /// </summary>
        /// <returns></returns>
        private SheetsService CreateConnectionGoogle() {
            GoogleCredential credential;

            // Чтение учетных данных из файла JSON
            using (var stream =
                new FileStream(json_path, FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(settings.Scopes);

            }

            // Создание сервиса Google Sheets с использованием учетных данных
            var service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }
        // internal string json_path = Path.GetFullPath(fileName);

    }
}

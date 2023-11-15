using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Google.Apis.Sheets.v4.SheetsService;
using PassportGenerator_Test.Model;
using OfficeOpenXml;
using System.Data;
using System.Windows.Forms;
using System.Configuration;

namespace PassportGenerator_Test {
    internal class ImportGoogle {
        public GoogleSettings settings = new GoogleSettings();
        private string json_path;
        private string fileName;
        private string spreadsheetId;
        private string ApplicationName;

        // internal string json_path = Path.GetFullPath(GoogleStg.fileName);
        // GoogleStg.ApplicationName = "PassportGenerator";
        public ImportGoogle(string spreadsheetId, string fileName) {
            settings.fileName = fileName;
            settings.spreadsheetId = spreadsheetId;
            settings.ApplicationName = "PassportGenerator";
            json_path = settings.json_path(fileName);
        }


        /// <summary>
        /// Запуск метода загрузки данных
        /// </summary>
        /// <param name="range"></param>
        /// <param name="excelfile_name"></param>
        internal void ImportDataFromGoogle(string range, string excelfile_name) {
            IList<IList<Object>> values = GetGoogleSheetsValue(json_path, spreadsheetId, range);
            FillInAnExcel(values, excelfile_name);
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

        /// <summary>
        /// Метод для заполнения файла Excel данными полученными из google таблицы
        /// </summary>
        /// <param name="values"></param>
        /// <param name="excefile_name"></param>
        private static void FillInAnExcel(IList<IList<Object>> values, string excefile_name) {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Если данные получены, они записываются в Excel
            if (values != null && values.Count > 0) {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excefile_name))) {

                    var worksheet = excelPackage.Workbook.Worksheets[0]; // Получить первый лист
                    int startRowInExcel = FindFirstEmptyRow(excelPackage, worksheet);

                    // Заполняем excel файл данными
                    for (int i = 0; i < values.Count; i++) {
                        for (int j = 0; j < values[i].Count; j++) {
                            worksheet.Cells[startRowInExcel + i, j + 1].Value = values[i][j];
                        }
                    }

                    excelPackage.Save();
                }
            }
        }

        /// <summary>
        /// Метод для поиска первое пустой строки
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="worksheet"></param>
        /// <param name="startRowInExcel"></param>
        static int FindFirstEmptyRow(ExcelPackage excelPackage, ExcelWorksheet worksheet) {
            int startRowInExcel = 1;
            while (worksheet.Cells[startRowInExcel, 1].Value != null && startRowInExcel <= worksheet.Dimension.End.Row) {
                startRowInExcel++;
            }
            return startRowInExcel;
        }

    }
}

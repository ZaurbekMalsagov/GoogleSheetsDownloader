using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace PassportGenerator_Test {
    public partial class Form1: Form {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "PassportGenerator";
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            string sheetlist_name = "Atom";
            string begin_column_address = "A";
            string begin_string_number = "3";
            string end_column_name = "J";
            string end_string_number = "21";
            //string excefile_name = "C:\\WORK\\PROJECTS\\PassportGenerator_Test\\test.xlsx";

            // Путь к файлу Excel, в который будут записаны данные
            string excefile_name = "C:\\WORK\\PROJECTS\\PassportGenerator_Test\\Устройства.xlsx";

            // Учетные данные Google для доступа к Google Sheets
            GoogleCredential credential;


            // Чтение учетных данных из файла JSON
            using (var stream =
                new FileStream("C:\\WORK\\PROJECTS\\PassportGenerator_Test\\passportgenerator-403707-030caf5e945d — копия.json", FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Создание сервиса Google Sheets с использованием учетных данных
            var service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // ID таблицы Google Sheets, из которой будут извлекаться данные
            // string spreadsheetId = "1CQa3-cxfpkUM5bzDUffMbXS0VmhFO3P868dL8Vt06Jg";
            string spreadsheetId =  "1Id_kNKmrMSUpalwY9rUf0uo8wNLBja9qo0Hw7Xe4gy0";
            
            // Формируем строку диапазона для извлечения данных
            string range = $"{sheetlist_name}!{begin_column_address}{begin_string_number}:{end_column_name}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Выполнение запроса и получение данных из Google Sheets
            IList<IList<Object>> values = request.Execute().Values;

            // Указания типа лицензии для библиотеки EPPlus, которая помогает работать с excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Если данные получены, они записываются в Excel
            if (values != null && values.Count > 0) {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excefile_name))) {
                    //var worksheet = excelPackage.Workbook.Worksheets.Add("S1");
                    var worksheet = excelPackage.Workbook.Worksheets[0]; // Получить первый лист
                    int startRowInExcel = FindFirstEmptyRow(excelPackage, worksheet);
                    
                    // Очищаем документ с конца и до второй строки
                    //ResetExcelRows(excelPackage, worksheet);

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
        /// Метод для связи с Google Sheets. Считывание учетных данных из json google cloud, определение ID таблицы из которой берутся данные. 
        /// </summary>
        /// <param name="range">Строка с диапазоном требуемых значений</param>
        /// <param name=""></param>
        /// <returns></returns>
        static IList<IList<Object>> ConnectGoogleSheets(string range) {
            GoogleCredential credential;


            // Чтение учетных данных из файла JSON
            using (var stream =
                new FileStream("C:\\WORK\\PROJECTS\\PassportGenerator_Test\\passportgenerator-403707-030caf5e945d — копия.json", FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Создание сервиса Google Sheets с использованием учетных данных
            var service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // ID таблицы Google Sheets, из которой будут извлекаться данные
            // string spreadsheetId = "1CQa3-cxfpkUM5bzDUffMbXS0VmhFO3P868dL8Vt06Jg";
            string spreadsheetId = "1Id_kNKmrMSUpalwY9rUf0uo8wNLBja9qo0Hw7Xe4gy0";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Выполнение запроса и получение данных из Google Sheets
            IList<IList<Object>> values = request.Execute().Values;

            return values;
        }

        /// <summary>
        /// Очищаем строки excel-файла начиная со второй строки
        /// </summary>
        /// <param name="excelPackage">Экземпляр ExcelPackage</param>
        /// <param name="worksheet_number">Номер листа</param>
        static void ResetExcelRows(ExcelPackage excelPackage, ExcelWorksheet worksheet) {
            
            int startRowInExcel = 2;
            for (int i = worksheet.Dimension.End.Row; i >= startRowInExcel; i--) {
                worksheet.DeleteRow(i);
            }
        }

        /// <summary>
        /// Метод для формирования строки диапазона считав данные с textBox формы
        /// </summary>
        /// <returns></returns>
        internal string GetRangeFromTxtBox() {
            string begin_row = textBox1.Text;
            string end_row = textBox2.Text;
            string sheetlist_name = textBox3.Text;
            string range = $"{sheetlist_name}!{begin_row}:{end_row}";

            return range;
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

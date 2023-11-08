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
using Newtonsoft.Json.Converters;
using static OfficeOpenXml.ExcelErrorValue;
using System.Runtime.CompilerServices;

namespace GoogleSheetsDownloader {
    public partial class Form1: Form {
        // Определяем права доступа пользователя 
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        // Указываем название проекта, который создали в Google Cloud
        static string ApplicationName = "PassportGenerator";
        
        public Form1() {
            InitializeComponent();
        }

        /// <summary>
        /// Действие на кнопку "Import data from Google"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e) {
            string json_path = ReadJson();
            string excefile_name = ExcelFileName();
            string range = GetRangeFromTxtBox();
            string spreadsheetId = GoogleSheetsID();


            IList<IList<Object>> values = ConnectGoogleSheets(json_path, spreadsheetId, range);
            FillInAnExcel(values, excefile_name);
        }

        /// <summary>
        /// Действие на кнопку "Очистить excel файл"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e) {
            string excefile_name = ExcelFileName();
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excefile_name))) {
                
                var worksheet = excelPackage.Workbook.Worksheets[0]; // Получить первый лист

                ResetExcelRows(excelPackage, worksheet);

                excelPackage.Save();
            }
            

        }

        /// <summary>
        /// Метод для связи с Google Sheets. Считывание учетных данных из json google cloud, определение ID таблицы из которой берутся данные. 
        /// </summary>
        /// <param name="range">Строка с диапазоном требуемых значений</param>
        /// <param name=""></param>
        /// <returns></returns>
        static IList<IList<Object>> ConnectGoogleSheets(string json_path, string spreadsheetId, string range) {
            GoogleCredential credential;


            // Чтение учетных данных из файла JSON
            using (var stream =
                new FileStream(json_path, FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);

            }

            // Создание сервиса Google Sheets с использованием учетных данных
            var service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
        private string GetRangeFromTxtBox() {
            string begin_row = txtBxStartRange.Text;
            string end_row = txtBxEndRange.Text;
            string sheetlist_name = txtBxListName.Text;

            // Формируем строку диапазона
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

        /// <summary>
        /// Метод для заполнения файла Excel данными полученными из google таблицы
        /// </summary>
        /// <param name="values"></param>
        /// <param name="excefile_name"></param>
        private void FillInAnExcel(IList<IList<Object>> values, string excefile_name) {
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
        /// Cчитывание ID Google таблицы
        /// </summary>
        /// <returns></returns>
        private string GoogleSheetsID() => txtBxIdTable.Text;
        
        /// <summary>
        /// Метод для считывания пути до excel-файла в который копируются данные
        /// </summary>
        /// <returns></returns>
        private string ExcelFileName() => txtBxExcelFile.Text;
        

        /// <summary>
        /// Метод для считывания пути до json-файла
        /// </summary>
        /// <returns></returns>
        private string ReadJson() => txtBxJsonPath.Text;
        
        
    }
}

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
using GoogleSheetsDownloader.Model;
using OfficeOpenXml;
using System.Data;
using System.Windows.Forms;
using System.Configuration;

namespace GoogleSheetsDownloader {
    internal class GoogleDataControl {
        public ImportDataFromGoogle settings = new ImportDataFromGoogle("");
        private string json_path;
        private string fileName;
        private string spreadsheetId;
        private string ApplicationName;

        // internal string json_path = Path.GetFullPath(GoogleStg.fileName);
        // GoogleStg.ApplicationName = "PassportGenerator";
        public GoogleDataControl(string spreadsheetId, string fileName) {
            settings.fileName = fileName;
            settings.spreadsheetId = spreadsheetId;
            settings.ApplicationName = "PassportGenerator";
            json_path = settings.json_path(fileName);
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
        /// Метод для поиска первое пустой строки в файле excel
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="worksheet"></param>
        /// <param name="startRowInExcel"></param>
        internal int FindFirstEmptyRow(ExcelPackage excelPackage, ExcelWorksheet worksheet) {
            int startRowInExcel = 1;
            while (worksheet.Cells[startRowInExcel, 1].Value != null && startRowInExcel <= worksheet.Dimension.End.Row) {
                startRowInExcel++;
            }
            return startRowInExcel;
        }

       

    }
}

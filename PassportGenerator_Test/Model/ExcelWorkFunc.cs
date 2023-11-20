using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsDownloader.Model {
    internal class ExcelWorkFunc {
        /// <summary>
        /// Метод для заполнения файла Excel данными полученными из google таблицы
        /// </summary>
        /// <param name="values"></param>
        /// <param name="excefile_name"></param>
        internal void FillInAnExcel(IList<IList<Object>> values, string excefile_name) {
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
        internal int FindFirstEmptyRow(ExcelPackage excelPackage, ExcelWorksheet worksheet) {
            int startRowInExcel = 1;
            while (worksheet.Cells[startRowInExcel, 1].Value != null && startRowInExcel <= worksheet.Dimension.End.Row) {
                startRowInExcel++;
            }
            return startRowInExcel;
        }

        /// <summary>
        /// Очищаем строки excel-файла начиная со второй строки
        /// </summary>
        /// <param name="excelPackage">Экземпляр ExcelPackage</param>
        /// <param name="worksheet_number">Номер листа</param>
        internal void ResetExcelRows(ExcelPackage excelPackage, ExcelWorksheet worksheet) {

            int startRowInExcel = 2;
            for (int i = worksheet.Dimension.End.Row; i >= startRowInExcel; i--) {
                worksheet.DeleteRow(i);
            }
        }
    }
}

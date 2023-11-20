using Google.Apis.Sheets.v4;
using GoogleSheetsDownloader.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsDownloader.Model {
    internal class GoogleSettings {

        /// <summary>
        /// Конструктор класса GoogleSettings в котором хранятся данные для связи с Google Sheets
        /// </summary>
        /// <param name="spreadsheetId">ID таблицы из которой берутся данные</param>
        /// <param name="fileName">Имя файла json</param>
        /// <param name="ApplicationName">Имя проекта, созданного в Google Cloud</param>
        public GoogleSettings(string spreadsheetId, string fileName, string ApplicationName) {
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



        // internal string json_path = Path.GetFullPath(fileName);

    }
}

using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsDownloader.Model {
    internal class GoogleSettings {
        // Определяем права доступа пользователя к таблице
        internal string[] Scopes { get; set; } = { SheetsService.Scope.SpreadsheetsReadonly };

        // Указываем название проекта, который создали в Google Cloud
        internal string ApplicationName { get; set; }

        public string fileName { get; set; } // "passportgenerator-403707-030caf5e945d.json";

        internal string spreadsheetId { get; set; }

        internal string json_path(string fileName) => Path.GetFullPath(fileName);



        // internal string json_path = Path.GetFullPath(fileName);

    }
}

using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PassportGenerator_Test.Model {
    internal class GoogleSettings {
        // Определяем права доступа пользователя к таблице
        internal string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        // Указываем название проекта, который создали в Google Cloud
        internal string ApplicationName = "PassportGenerator";

        internal string fileName = "passportgenerator-403707-030caf5e945d.json";

        // internal string json_path = Path.GetFullPath(fileName);
    }
}

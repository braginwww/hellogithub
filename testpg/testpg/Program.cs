using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Npgsql;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Timers;

namespace testpg
{
    class Program
    {
        static string cspath = Directory.GetCurrentDirectory() + @"\constrings.txt"; // файл где хранится строки подключения к бд
        static string gapath = Directory.GetCurrentDirectory() + @"\googleaccountfile.txt"; //googleaccount values
        //static string cspath = @"c:\temp\constrings.txt";
        //static string gapath = @"c:\temp\googleaccountfile.txt";
        internal static int timerLength = 5000; // время таймера

        static string _myGoogleAccount;
        static string _myApplicationName;
        static string _mySpreadsheetId;
        static string[] _myScopes = { SheetsService.Scope.Spreadsheets };

        static MyDbInfo _myDbInfo; // класс который хранит имя сервера, имя бд и строку подключения к бд
        static List<MyDbInfo> _myDbInfoList = new List<MyDbInfo>(); // лист на основе класса MyDbInfo

        static SheetsService _myService;


        // событие которое срабатывает после завешения таймера
        /// <summary>
        /// ///////////таймер///таймер/////////таймер////////////таймер/////////////////////таймер///////////////////////
        /// </summary>
        private static void OnTimedEvent(Object source, ElapsedEventArgs e, string path)
        {
            foreach (MyDbInfo item in Program._myDbInfoList)
            {
                // заносим данные в лист с именем сервера
                _myUpdateSpreadSheet(item.ServerName, item.DataBaseName, _myCalcDbSize(item.ConString, item.DataBaseName));
            }
        }

        /// <summary>
        /// //////////main////////////main////////////////////main///////////////////////main///////////////main//////////
        /// </summary>
        static void Main(string[] args)
        {
            System.Timers.Timer aTimer = new System.Timers.Timer(timerLength); // создаем таймер для повторений
            aTimer.Elapsed += (sender, e) => OnTimedEvent(sender, e, cspath);
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
            
            _myReadGoogleAccountFile(); // читаем data из файла Google Account
            _myCreateDocument(); // создает файл в гугл таблицах
            _myReadConStringsFile(); // считываем данные из файла с connetctionstrings.txt и заносим в лист _myDbInfoList
            _myCreteSheetsWithServerName(); // создаем листы (имена листов берутся из _myDbInfoList из поля ServerName)
            Console.ReadKey();
        }
        static void _myReadGoogleAccountFile()
        {
            using (StreamReader sr = new StreamReader(gapath))
            {
                string gaString;

                // передаем значение строки файла в переменную conString
                while ((gaString = sr.ReadLine()) != null)
                {
                    string[] ga = gaString.Split(new char[] { ';' });
                    string[] acc = ga[0].Split(new char[] { '=' });
                    string[] app = ga[1].Split(new char[] { '=' });
                    string[] sid = ga[2].Split(new char[] { '=' });
                    _myGoogleAccount = acc[1]; // глобальная переменная
                    _myApplicationName = app[1]; // глобальная переменная
                    _mySpreadsheetId = sid[1]; // глобальная переменная
                }
            }
        }
        static void _myCreateDocument() // создает файл в гугл таблицах
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    _myScopes,
                    _myGoogleAccount,
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            _myService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _myApplicationName,
            });
        }

        static void _myCreateSheet(string sheetName) // метод создающий лист с заданным именем sheetName
                                                     // он нам нам понадобится в _myCreteSheetsWithServerName
                                                     // когда будут создаваться лист с именем как у Server

        {
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = sheetName;
            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });

            var batchUpdateRequest = _myService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, _mySpreadsheetId);
            batchUpdateRequest.Execute();
        }

        static void _myUpdateSpreadSheet(string sheetName, string DbName, string size) // заносим в лист значения
                                                                                       // этот метод запускается в таймере
        {
            var range = $"{sheetName}!A:D";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { sheetName, DbName, size, DateTime.Now.ToString("dd.MM.yyyy") };
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = _myService.Spreadsheets.Values.Append(
                valueRange,
                _mySpreadsheetId,
                range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        //читаем файл со строками подключения к БД и заносим данные в класс
        static void _myReadConStringsFile()
        {
            // читаем файл constrings.txt и создаем лист с информацией о среверах
            using (StreamReader sr = new StreamReader(cspath))
            {
                string conString;

                // передаем значение строки файла в переменную conString
                while ((conString = sr.ReadLine()) != null)
                {
                    string[] splittedConString = conString.Split(new char[] { ';' });
                    string[] srv = splittedConString[0].Split(new char[] { '=' });
                    string[] db = splittedConString[4].Split(new char[] { '=' });
                    var _myDbInfo = new MyDbInfo(); // класс который хранит имя сервера, имя бд и строку подключения к бд
                    _myDbInfo.ServerName = srv[1];
                    _myDbInfo.DataBaseName = db[1];
                    _myDbInfo.ConString = conString;
                    Console.WriteLine(_myDbInfo.ServerName);
                    _myDbInfoList.Add(_myDbInfo);

                }
            }
        }
        static void _myCreteSheetsWithServerName() // создаем листы с именем сервера из цикла; метод запускается один раз в Main
        {
            foreach (MyDbInfo item in _myDbInfoList)
            {
                try
                {
                    _myCreateSheet(item.ServerName); // создаем лист с именем item.ServerName
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Лист с именем {item.ServerName} уже существует");
                }
            }
        }
        // считает размер БД в Гб
        static string _myCalcDbSize(string connStr, string dbName)
        {
            NpgsqlConnection m_conn = new NpgsqlConnection(connStr);
            NpgsqlCommand query_cmd = new NpgsqlCommand($"SELECT pg_database_size('{dbName}');", m_conn);
            try
            {
                m_conn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(dbName + " does not exists;");
                return "error";
            }

            float size = (Int64)query_cmd.ExecuteScalar() / 1024 / 1024 / 1024; // ГБ
            Console.WriteLine(size.ToString("F"));
            m_conn.Close();
            return size.ToString("F");
        }
    }
}

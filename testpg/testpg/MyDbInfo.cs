namespace testpg
{
    class MyDbInfo // используется в методе _myReadConStringsFile()
    {
        public string ServerName { get; set; } // Хранит имя сервера
        public string DataBaseName { get; set; } // Хранит имя базы данных
        public string ConString { get; set; } // Хранит строку подключения к БД
    }
}

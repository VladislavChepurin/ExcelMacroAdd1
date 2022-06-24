using System;
using System.Data.Linq;
using System.Data.OleDb;
using System.Windows.Forms;
using ExcelMacroAdd.DBEntity;
using System.Data.Linq.Mapping;


namespace ExcelMacroAdd
{

    /// <summary>
    /// Класс доступа к базе данных
    /// </summary>
    public class DBConect
    {
        public struct DBtable
        {
            public string IpTable { get; set; }
            public string KlimaTable { get; set; }
            public string ReserveTable { get; set; }
            public string HeightTable { get; set; }
            public string WidthTable { get; set; }
            public string DepthTable { get; set; }
            public string ArticleTable { get; set; }
            public string ExecutionTable { get; set; }
            public string VendorTable { get; set; }
        }
        // Переменная подключения к БД - static
        private static OleDbConnection myConnection;

        private DataContext db;

        // Путь к базе данных
#if DEBUG
        private readonly string _pPatch = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Прайсы\Макро\";
#else
        private readonly string _pPatch = @"\\192.168.100.100\ftp\Info_A\FTP\Производство Абиэлт\Инженеры\"; // Путь к базе данных
#endif
        private readonly string _sPatch = "BdMacro.accdb";

        private readonly string _providerData = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=";

        public string PPatch {get;}

        public DBConect()
        {
            PPatch = _pPatch;
        }

        /// <summary>
        /// Отрытие соединения с базой данных
        /// </summary>
        public void OpenDB()
        {        
            try
            {
                db = new DataContext(_pPatch + _sPatch );

                myConnection = new OleDbConnection(_providerData + _pPatch + _sPatch + ";");
                // открываем соединение с БД
                myConnection.Open();
            }
            catch (OleDbException)
            {
                MessageBox.Show(
                "База данных не найдена, убедитесь в наличии файла базы данных и сетевого подключения. " +
                "Файл " + _pPatch + _sPatch + " не найден в предпологаемом местонахождении.",
                "Ошибка базы данных",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        /// <summary>
        /// Закрытие соединения с БД
        /// </summary>
        public void CloseDB()
        {
         //   db.Dispose();


            myConnection.Dispose();
            myConnection.Close();
        }

        /// <summary>
        /// Запрашавает одно значение из базы данных
        /// </summary>
        /// <param name="_requestDB"></param>
        /// <returns></returns>
        public string RequestDB(string requestDB, int colum)     // Считывание одного значения из базы данных
        {
            try
            {

          
                Table<Settings> settings = db.GetTable<Settings>();

                if (!(settings == null))
                {
                    foreach (var item in settings)
                    {
                        Console.WriteLine(item.SetName +"   "  + item.SetOption);
                    }
                }
             /*
                var result = from u in db.GetTable<Settings>()
                             where u.SetName == requestDB
                                select u;

                */


                string rt = default;
                // Собираем запрос к БД
                OleDbCommand command = new OleDbCommand(requestDB, myConnection);

                OleDbDataReader reader = command.ExecuteReader();
                // Считываем и возвращаем значение их базы данных
                while (reader.Read())
                {
                    rt = (reader[colum].ToString());
                }
                return rt;
            }
            catch (OleDbException exception)
            {
                Message(exception);
                return default;
            }       
            catch (InvalidOperationException) 
            {
                return default;
            }

        }

        /// <summary>
        /// Запрашивает наличие записи в базе данных
        /// </summary>
        /// <param name="dataRequest"></param>
        /// <returns></returns>
        public Boolean CheckReadDB(string dataRequest)
        {
            try
            {
                OleDbCommand commandRead = new OleDbCommand(dataRequest, myConnection);
                if (commandRead.ExecuteScalar() == null)
                {
                    commandRead.Dispose();
                    return true;
                }
                else
                {
                    commandRead.Dispose();
                    return false;
                }
            }
            catch (OleDbException exception)
            {
                Message(exception);
                return false;
            }           
        }

        /// <summary>
        /// Отправляет запрос в БД без возрращаемого значения (UPDATE)
        /// </summary>
        /// <param name="queryUpdate"></param>
        /// <param name="dataRequest"></param>
        public void MetodDB(string queryUpdate, string dataRequest)
        {
            try
            {
                OleDbCommand commandUpdate = new OleDbCommand(queryUpdate, myConnection)
                {
                    Connection = myConnection,
                    // Строка запроса SQL
                    CommandText = dataRequest
                };
                commandUpdate.ExecuteNonQuery();
                // Освобождаем процессы
                commandUpdate.Dispose();
            }
            catch (OleDbException exception)
            {
                Message(exception);
            }     
        }
        /// <summary>
        /// Пишет в структуру передан по референсной ссылке данные из базы данных
        /// </summary>
        /// <param name="dataRead"></param>
        /// <param name="dBtable"></param>
        public DBtable ReadingDB(string dataRead)
        {
            try
            {
                DBtable dBtable = default;
                OleDbCommand command = new OleDbCommand(dataRead, myConnection);
                OleDbDataReader reader = command.ExecuteReader();
                // Чтение из базы данных и поэлементная запись в массив
                while (reader.Read())
                {
                    dBtable.IpTable        = reader[1].ToString();
                    dBtable.HeightTable    = reader[2].ToString();
                    dBtable.KlimaTable     = reader[3].ToString();
                    dBtable.ReserveTable   = reader[4].ToString();
                    dBtable.WidthTable     = reader[5].ToString();
                    dBtable.DepthTable     = reader[6].ToString();
                    dBtable.ArticleTable   = reader[7].ToString();
                    dBtable.ExecutionTable = reader[8].ToString();
                    dBtable.VendorTable    = reader[9].ToString();
                }
                return dBtable;
            }
            catch (OleDbException exception)
            {                
                Message(exception);
                return default;
            }
        }

        /// <summary>
        /// Местный метод вывода ошибок
        /// </summary>
        /// <param name="exception"></param>
        private void Message(Exception exception)
        {
            MessageBox.Show(
            exception.ToString(),
            "Ошибка базы данных",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error,
            MessageBoxDefaultButton.Button1,
            MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}

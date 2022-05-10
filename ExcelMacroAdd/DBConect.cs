using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    interface IDBConect{

        void OpenDB();
        void CloseDB();
        string RequestDB(string requestDB, int colum);
        Boolean CheckReadDB(string dataRequest);
        void MetodDB(string queryUpdate, string dataRequest);
        void ReadingDB(string dataRead, ref DBtable dBtable);
    }


    public struct DBtable 
    {
        public string ipTable { get; set; }
        public string klimaTable { get; set; }
        public string reserveTable { get; set; }
        public string heightTable { get; set; }
        public string widthTable { get; set; }
        public string depthTable { get; set; }
        public string articleTable { get; set; }
        public string executionTable { get; set; }
        public string vendorTable { get; set; }
    }

    /// <summary>
    /// Класс доступа к базе данных
    /// </summary>
    internal class DBConect: IDBConect
    {
        // Переменная подключения к БД - static
        private static OleDbConnection myConnection;

        // Путь к базе данных
#if DEBUG
        private readonly string _pPatch = @"C:\Users\ПК\Desktop\Прайсы\Макро\";
#else
        private readonly string _pPatch = @"\\192.168.100.100\ftp\Info_A\FTP\Производство Абиэлт\Инженеры\"; // Путь к базе данных
#endif
        private readonly string _sPatch = "BdMacro.mdb";

        private readonly string _providerData = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="; 
        
        public DBConect()
        {     
            pPatch = _pPatch;
            sPatch = _sPatch;
        }
                   
        public string pPatch { get; }
        public string sPatch { get; }

        /// <summary>
        /// Отрытие соединения с базой данных
        /// </summary>
        public void OpenDB()
        {
            myConnection = new OleDbConnection(_providerData + _pPatch + _sPatch + ";");
            // открываем соединение с БД
            myConnection.Open();
        }
        
        /// <summary>
        /// Закрытие соединения с БД
        /// </summary>
        public void CloseDB()
        {
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
                return null;
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
        public void ReadingDB(string dataRead, ref DBtable dBtable)
        {
            try
            {
                OleDbCommand command = new OleDbCommand(dataRead, myConnection);
                OleDbDataReader reader = command.ExecuteReader();
                // Чтение из базы данных и поэлементная запись в массив
                while (reader.Read())
                {
                    dBtable.ipTable        = reader[1].ToString();
                    dBtable.heightTable    = reader[2].ToString();
                    dBtable.klimaTable     = reader[3].ToString();
                    dBtable.reserveTable   = reader[4].ToString();
                    dBtable.widthTable     = reader[5].ToString();
                    dBtable.depthTable     = reader[6].ToString();
                    dBtable.articleTable   = reader[7].ToString();
                    dBtable.executionTable = reader[8].ToString();
                    dBtable.vendorTable    = reader[9].ToString();
                }
            }
            catch (OleDbException exception)
            {
                Message(exception);
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

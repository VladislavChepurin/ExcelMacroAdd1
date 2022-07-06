using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.UserVariables;
using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace ExcelMacroAdd.Servises
{   
    /// <summary>
    /// Класс доступа к базе данных
    /// </summary>
    public class DBConect: IDBConect
    {     
        // Переменная подключения к БД - static
        private static OleDbConnection myConnection;

        // Путь к базе данных
#if DEBUG
        private readonly string _pPatch = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"Прайсы\Макро\");
#else
        private readonly string _pPatch = @"\\192.168.100.100\ftp\Info_A\FTP\Производство Абиэлт\Инженеры\"; // Путь к базе данных
#endif
        private readonly string _sPatch = "BdMacro.accdb";

        private readonly string _providerData = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=";


        public DBConect()
        {
            PPatch = _pPatch;
        }

        public string PPatch { get;}

        /// <summary>
        /// Отрытие соединения с базой данных
        /// </summary>
        public void OpenDB()
        {
            try
            {
                myConnection = new OleDbConnection(_providerData + Path.Combine(_pPatch, _sPatch) + ";");
                // открываем соединение с БД
                myConnection.Open();
            }
            catch (OleDbException)
            {
                MessageBox.Show(
                "База данных не найдена, убедитесь в наличии файла базы данных и сетевого подключения. " +
                "Файл " + Path.Combine(_pPatch, _sPatch).ToString() + " не найден в предпологаемом местонахождении.",
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
            myConnection.Dispose();
            myConnection.Close();
        }

        /// <summary>
        /// Запрашавает одно значение из базы данных
        /// </summary>
        /// <param name="_requestDB"></param>
        /// <returns></returns>
        public string ReadOnlyOneNoteDB(string requestDB, int colum)     // Считывание одного значения из базы данных
        {
            try
            {
                string rt = default;
                // Собираем запрос к БД
                OleDbCommand command = new OleDbCommand(requestDB, myConnection);
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    // Считываем и возвращаем значение их базы данных
                    while (reader.Read())
                    {
                        rt = (reader[colum].ToString());
                    }
                    return rt;
                }
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
        /// Отправляет запрос в БД без возрращаемого значения (UPDATE)
        /// </summary>
        /// <param name="queryUpdate"></param>
        /// <param name="dataRequest"></param>
        public void UpdateNotesDB(string queryUpdate, string dataRequest)
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
        public DBtable ReadSeveralNotesDB(string dataRead)
        {
            try
            {
                DBtable dBtable = new DBtable();
                OleDbCommand command = new OleDbCommand(dataRead, myConnection);
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    // Чтение из базы данных и поэлементная запись в массив
                    while (reader.Read())
                    {
                        dBtable.IpTable = reader[1].ToString();
                        dBtable.HeightTable = reader[2].ToString();
                        dBtable.KlimaTable = reader[3].ToString();
                        dBtable.ReserveTable = reader[4].ToString();
                        dBtable.WidthTable = reader[5].ToString();
                        dBtable.DepthTable = reader[6].ToString();
                        dBtable.ArticleTable = reader[7].ToString();
                        dBtable.ExecutionTable = reader[8].ToString();
                        dBtable.VendorTable = reader[9].ToString();
                    }
                    return dBtable;
                }
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

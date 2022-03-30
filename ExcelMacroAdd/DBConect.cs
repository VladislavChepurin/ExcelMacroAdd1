using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    /// <summary>
    /// Класс доступа к базе данных
    /// </summary>
    public class DBConect
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
        /// Запрашавает настройки в базе данных
        /// </summary>
        /// <param name="_requestDB"></param>
        /// <returns></returns>
        public string RequestDB(string requestDB)     // Считывание настроек в базе данных 
        {
            try
            {
                string rt = null;
                // Собираем запрос к БД
                OleDbCommand command = new OleDbCommand(requestDB, myConnection);

                OleDbDataReader reader = command.ExecuteReader();
                // Считываем и возвращаем значение их базы данных
                while (reader.Read())
                {
                    rt = (reader[2].ToString());
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
        /// Пишет в массив переданый по референсной ссылке данные из базы данных
        /// </summary>
        /// <param name="dataRead"></param>
        /// <param name="dataMassiv"></param>
        public void ReadingDB(string dataRead, ref string[] dataMassiv)
        {
            try
            {
                OleDbCommand command = new OleDbCommand(dataRead, myConnection);
                OleDbDataReader reader = command.ExecuteReader();
                // Чтение из базы данных и поэлементная запись в массив
                while (reader.Read())
                {
                    dataMassiv[0] = reader[0].ToString();
                    dataMassiv[1] = reader[1].ToString();
                    dataMassiv[2] = reader[2].ToString();
                    dataMassiv[3] = reader[3].ToString();
                    dataMassiv[4] = reader[4].ToString();
                    dataMassiv[5] = reader[5].ToString();
                    dataMassiv[6] = reader[6].ToString();
                    dataMassiv[7] = reader[7].ToString();
                    dataMassiv[8] = reader[8].ToString();
                    dataMassiv[9] = reader[9].ToString();
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

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
    internal class DBConect: IDBConect
    {
        private static DBConect Connection;
        // Переменная подключения к БД - static
        private static OleDbConnection myConnection;
        // Путь к базе данных

        private readonly string PPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");     
        private readonly IResourcesDBConect resources;

        public DBConect(IResourcesDBConect resources)
        {
            this.resources = resources;    
        }              

        /// <summary>
        /// Отрытие соединения с базой данных
        /// </summary>
        public void OpenDB()
        {
            try
            {
                myConnection = new OleDbConnection(resources.ProviderData + Path.Combine(PPath, resources.NameFileDB) + ";");
                // открываем соединение с БД
                myConnection.Open();
            }
            catch (OleDbException)
            {
                MessageBox.Show(
                "База данных не найдена, убедитесь в наличии файла базы данных и сетевого подключения. " +
                "Файл " + Path.Combine(PPath, resources.NameFileDB).ToString() + " не найден в предпологаемом местонахождении.",
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
        /// Отправляет запрос в БД без возрращаемого значения Запись/обновление
        /// </summary>
        /// <param name="dataRequest"></param>
        public void UpdateNotesDB(string dataRequest)
        {
            try
            {



                using (OleDbCommand commandUpdate = new OleDbCommand(dataRequest, myConnection))
                {
                    commandUpdate.ExecuteNonQuery();
                }   
                




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
        
        internal static DBConect GetConnectionInstance(IResourcesDBConect resourcesDB)
        {
            if (Connection == null)
                Connection = new DBConect(resourcesDB);
            return Connection;
        }
    }
}

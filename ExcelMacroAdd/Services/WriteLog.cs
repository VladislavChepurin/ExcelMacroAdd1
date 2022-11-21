using System;
using System.IO;

namespace ExcelMacroAdd.Forms.Services
{
    internal static class WriteLog
    {
        /// <summary>
        /// Метод для записи логов шапки документа
        /// </summary>
        /// <param name="folder"></param>
        internal static void Logger(string folder)
        {
            string patch = Path.Combine(folder, "log.txt");
            using (StreamWriter output = File.AppendText(patch))
            {
                output.WriteLine("Версия OC:          " + Environment.OSVersion);
                output.WriteLine("Имя пользователя:   " + Environment.UserName);
                output.WriteLine("Имя компьютера:     " + Environment.MachineName);
                output.WriteLine("--------------------------------------------------------------------------------");
            }
        }

        /// <summary>
        /// Метод для записи логов формиррования паспортов
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="saveNum"></param> 
        /// <param name="amount"></param>
        /// <param name="verify"></param>
        internal static void Logger(string folder, string saveNum, int amount)
        {
            string patch = Path.Combine(folder, "log.txt");
            using (StreamWriter output = File.AppendText(patch))
            {
                output.WriteLine($"{DateTime.Now} | Паспорт {saveNum} сформирован успешно, в паспорте {amount} листа");
            }
        }
    }
}

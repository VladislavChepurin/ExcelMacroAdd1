using System;
using System.IO;
using System.Threading;

namespace ExcelMacroAdd.Services
{
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logDirectory = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "Logs"
        );
                
        public enum LogLevel
        {
            Debug,
            Info,
            Warning,
            Error,
            Critical
        }

        static Logger()
        {
            // Создаем директорию для логов при первом использовании
            if (!Directory.Exists(_logDirectory))
            {
                Directory.CreateDirectory(_logDirectory);
            }
        }

        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            try
            {
                Monitor.Enter(_lock);
                string logFile = Path.Combine(
                    _logDirectory,
                    $"log-{DateTime.Today:yyyy-MM-dd}.txt"
                );

                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} " +
                                  $"[{level.ToString().ToUpper()}] " +
                                  $"{message}";

                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Резервное логирование в консоль при ошибках записи
                Console.WriteLine($"Logger failed: {ex.Message}");
            }
            finally
            {
                Monitor.Exit(_lock);
            }
        }

        public static void LogException(Exception ex, string context = null)
        {
            string message = $"EXCEPTION: {ex.GetType().Name}\n" +
                           $"Message: {ex.Message}\n" +
                           $"Stack Trace: {ex.StackTrace}";

            if (!string.IsNullOrEmpty(context))
            {
                message = $"[{context}] " + message;
            }

            Log(message, LogLevel.Error);
        }
    }
}
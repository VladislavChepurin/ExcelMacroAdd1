using ExcelMacroAdd.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;


namespace ExcelMacroAdd.Functions
{
    internal sealed class AutoCadCalled : AbstractFunctions
    {
        private const int AtrticleColumn = 1;

        //public AutoCadCalled()
        //{

        //}

        public override async void Start()
        {
            string exePath = "AutoCadLancher.exe";
            //if (!File.Exists(exePath))
            //{
            //    MessageError($"Ошибка: Файл {exePath} не найден!", "Ошибка");
            //    return;
            //}

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow - 1;
            var currentRow = firstRow;

            var blocks = new List<string>();

            do
            {
                try
                {
                    string sArticle = Convert.ToString(Worksheet.Cells[currentRow, AtrticleColumn].Value2);

                    if (!String.IsNullOrEmpty(sArticle))
                    {
                        blocks.Add(sArticle);
                    }
                 }
               
                catch (Exception ex)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    return;
                }
                currentRow++;
            }
            while (currentRow <= endRow);
                        
            var arguments = new List<string>()
            {
                $"-source \"drawing.dwg\"",
                $"-blocks {string.Join(" ", blocks)}",
                $"-output \"output.dwg\""
            };
            await RunProcess(exePath, string.Join(" ", arguments));
        }

        private async Task RunProcess(string exePath, string arguments)
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = false
                };

                using (var process = new Process())
                {
                    process.StartInfo = processInfo;

                    // Обработчики событий для вывода в реальном времени
                    //process.OutputDataReceived += (sender, e) =>
                    //{
                    //    if (!string.IsNullOrEmpty(e.Data))
                    //        Console.WriteLine(e.Data);
                    //};

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                            MessageError($"ОШИБКА: {e.Data}", "Ошибка");                          
                    };
                                       
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    // Ждем завершения процесса
                    await Task.Run(() => process.WaitForExit());                    

                    if (process.ExitCode == 0)
                    {
                        MessageInformation("✅ Операция выполнена успешно!", "Информация");                 
                    }
                    else
                    {
                        MessageError($"Операция завершена с ошибками!", "Ошибка");                       
                    }
                }
            }
            catch (Exception ex)
            {
                MessageError($"💥 Ошибка при запуске процесса: {ex.Message}", "Ошибка");
                Logger.LogException(ex);
            }
        }
    }
}

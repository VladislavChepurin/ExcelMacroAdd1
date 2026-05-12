using System;
using System.IO;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.DataLayer.Infrastructure
{
    public sealed class AppContextFactory : IAppContextFactory
    {
        private const string DatabaseFileName = "BdMain.sqlite";
        private readonly string _databaseFilePath;

        public AppContextFactory(string databasePathOrDirectory)
        {
            if (string.IsNullOrWhiteSpace(databasePathOrDirectory))
            {
                throw new ArgumentException("Путь к базе данных не может быть пустым.", nameof(databasePathOrDirectory));
            }

            _databaseFilePath = ResolveDatabaseFilePath(databasePathOrDirectory);
        }

        public AppContext Create()
        {
            return new AppContext(_databaseFilePath);
        }

        private static string ResolveDatabaseFilePath(string databasePathOrDirectory)
        {
            string fullPath = Path.GetFullPath(databasePathOrDirectory);
            return fullPath.EndsWith(".sqlite", StringComparison.OrdinalIgnoreCase)
                ? fullPath
                : Path.Combine(fullPath, DatabaseFileName);
        }
    }
}

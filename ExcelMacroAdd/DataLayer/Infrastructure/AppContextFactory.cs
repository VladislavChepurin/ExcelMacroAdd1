using System;
using System.IO;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.DataLayer.Infrastructure
{
    public sealed class AppContextFactory : IAppContextFactory
    {
        private readonly string _dataDirectoryPath;

        public AppContextFactory(string dataDirectoryPath)
        {
            if (string.IsNullOrWhiteSpace(dataDirectoryPath))
            {
                throw new ArgumentException("Путь к каталогу базы данных не может быть пустым.", nameof(dataDirectoryPath));
            }

            _dataDirectoryPath = Path.GetFullPath(dataDirectoryPath);
        }

        public AppContext Create()
        {
            return new AppContext(_dataDirectoryPath);
        }
    }
}

using ExcelMacroAdd.BusinessLayer;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Infrastructure;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Data.Entity;
using System.IO;
using System.Linq;

namespace ExelMacroAdd.Tests
{
    internal sealed class SqliteIntegrationTestScope : IDisposable
    {
        private readonly string tempDirectory;
        private bool disposed;

        static SqliteIntegrationTestScope()
        {
            Database.SetInitializer<ExcelMacroAdd.DataLayer.Entity.AppContext>(null);
        }

        public SqliteIntegrationTestScope()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "ExcelMacroAddTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
            AppDomain.CurrentDomain.SetData("DataDirectory", tempDirectory + Path.DirectorySeparatorChar);

            string sourceDatabase = Path.Combine(GetSolutionRoot(), "ExcelMacroAdd", "DataLayer", "DataBase", "BdMain.sqlite");
            string targetDatabase = Path.Combine(tempDirectory, "BdMain.sqlite");
            File.Copy(sourceDatabase, targetDatabase, overwrite: false);

            MemoryCache = new MemoryCache(new MemoryCacheOptions());
            var appContextFactory = new AppContextFactory(tempDirectory + Path.DirectorySeparatorChar);
            UnitOfWorkFactory = new UnitOfWorkFactory(appContextFactory);
        }

        public IMemoryCache MemoryCache { get; }

        public IUnitOfWorkFactory UnitOfWorkFactory { get; }

        public JournalNkuQueryService CreateJournalQueryService()
        {
            return new JournalNkuQueryService(UnitOfWorkFactory, MemoryCache);
        }

        public JournalNkuWriteService CreateJournalWriteService()
        {
            return new JournalNkuWriteService(UnitOfWorkFactory, MemoryCache);
        }

        public NotPriceComponentsService CreateNotPriceComponentsService()
        {
            return new NotPriceComponentsService(UnitOfWorkFactory);
        }

        public (string MaterialName, int MaterialId) GetFirstMaterial()
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                var material = unitOfWork.Context.Materials
                    .OrderBy(x => x.Id)
                    .Select(x => new { x.Id, x.MaterialValue })
                    .First();

                return (material.MaterialValue, material.Id);
            }
        }

        public (string ExecutionName, int ExecutionId) GetFirstExecution()
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                var execution = unitOfWork.Context.Executions
                    .OrderBy(x => x.Id)
                    .Select(x => new { x.Id, x.ExecutionValue })
                    .First();

                return (execution.ExecutionValue, execution.Id);
            }
        }

        public string GetFirstMultiplicityName()
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                return unitOfWork.Context.Multiplicities
                    .OrderBy(x => x.Id)
                    .Select(x => x.Value)
                    .FirstOrDefault();
            }
        }

        public BoxBase GetJournalRecord(string article)
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                return unitOfWork.Context.JornalNkus
                    .FirstOrDefault(x => x.Article == article);
            }
        }

        public NotPriceComponent GetNotPriceComponentRecord(string article)
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                return unitOfWork.Context.NotPriceComponents
                    .Include("ProductVendor")
                    .Include("Multiplicity")
                    .FirstOrDefault(x => x.Article == article);
            }
        }

        public ProductVendor GetProductVendor(string vendorName)
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                return unitOfWork.Context.ProductVendors
                    .FirstOrDefault(x => x.VendorName == vendorName);
            }
        }

        public void SeedJournalRecord(BoxBase entity)
        {
            using (var unitOfWork = UnitOfWorkFactory.Create())
            {
                unitOfWork.Context.JornalNkus.Add(entity);
                unitOfWork.SaveChangesAsync().GetAwaiter().GetResult();
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            (MemoryCache as IDisposable)?.Dispose();

            if (Directory.Exists(tempDirectory))
            {
                try
                {
                    Directory.Delete(tempDirectory, recursive: true);
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        private static string GetSolutionRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null && !File.Exists(Path.Combine(current.FullName, "ExcelMacroAdd.sln")))
            {
                current = current.Parent;
            }

            if (current == null)
            {
                throw new DirectoryNotFoundException("Не удалось найти корень решения ExcelMacroAdd.sln.");
            }

            return current.FullName;
        }
    }
}

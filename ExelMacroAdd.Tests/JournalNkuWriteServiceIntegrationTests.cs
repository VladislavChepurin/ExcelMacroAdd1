using ExcelMacroAdd.BusinessLayer;
using ExcelMacroAdd.BusinessLayer.Models;
using ExcelMacroAdd.DataLayer.Entity;
using NUnit.Framework;
using System.Threading.Tasks;

namespace ExelMacroAdd.Tests
{
    [TestFixture]
    public class JournalNkuWriteServiceIntegrationTests
    {
        [Test]
        public async Task AddBoxAsync_PersistsRecord_AndInvalidatesMissingArticleCache()
        {
            using (var scope = new SqliteIntegrationTestScope())
            {
                var queryService = scope.CreateJournalQueryService();
                var writeService = scope.CreateJournalWriteService();
                var material = scope.GetFirstMaterial();
                var execution = scope.GetFirstExecution();
                string article = "TEST-BOX-" + TestContext.CurrentContext.Test.ID;

                var cachedMiss = await queryService.GetEntityJournal(article);
                Assert.IsNull(cachedMiss);

                var result = await writeService.AddBoxAsync(new JournalNkuWriteRequest
                {
                    Article = article,
                    Ip = 54,
                    Climate = "УХЛ1",
                    Weight = "123",
                    Height = "2000",
                    Width = "600",
                    Depth = "400",
                    MaterialName = material.MaterialName,
                    ExecutionName = execution.ExecutionName
                });

                Assert.AreEqual(JournalNkuWriteStatus.Added, result.Status);

                var recordFromCacheAwareQuery = await queryService.GetEntityJournal(article);
                var recordFromDatabase = scope.GetJournalRecord(article.ToLowerInvariant());

                Assert.Multiple(() =>
                {
                    Assert.IsNotNull(recordFromCacheAwareQuery);
                    Assert.IsNotNull(recordFromDatabase);
                    Assert.AreEqual("2000", recordFromDatabase.Height);
                    Assert.AreEqual("600", recordFromDatabase.Width);
                    Assert.AreEqual("400", recordFromDatabase.Depth);
                    Assert.AreEqual(material.MaterialId, recordFromDatabase.MaterialBoxId);
                    Assert.AreEqual(execution.ExecutionId, recordFromDatabase.ExecutionBoxId);
                    Assert.AreEqual("2000", recordFromCacheAwareQuery.Height);
                });
            }
        }

        [Test]
        public async Task UpdateBoxAsync_PersistsChanges_AndInvalidatesCachedHit()
        {
            using (var scope = new SqliteIntegrationTestScope())
            {
                var queryService = scope.CreateJournalQueryService();
                var writeService = scope.CreateJournalWriteService();
                var material = scope.GetFirstMaterial();
                var execution = scope.GetFirstExecution();
                int productVendorId = scope.GetFirstProductVendorId();
                string article = ("TEST-UPDATE-" + TestContext.CurrentContext.Test.ID).ToLowerInvariant();

                scope.SeedJournalRecord(new BoxBase
                {
                    Article = article,
                    Ip = 31,
                    Climate = "У2",
                    Weight = "50",
                    Height = "1800",
                    Width = "500",
                    Depth = "300",
                    MaterialBoxId = material.MaterialId,
                    ProductVendorId = productVendorId,
                    ExecutionBoxId = execution.ExecutionId
                });

                var cachedHit = await queryService.GetEntityJournal(article);
                Assert.IsNotNull(cachedHit);
                Assert.AreEqual("1800", cachedHit.Height);

                var result = await writeService.UpdateBoxAsync(new JournalNkuWriteRequest
                {
                    Article = article,
                    Ip = 65,
                    Climate = "У1",
                    Weight = "75",
                    Height = "2200",
                    Width = "800",
                    Depth = "500",
                    MaterialName = material.MaterialName,
                    ExecutionName = execution.ExecutionName
                });

                Assert.AreEqual(JournalNkuWriteStatus.Updated, result.Status);

                var updatedFromQuery = await queryService.GetEntityJournal(article);
                var updatedFromDatabase = scope.GetJournalRecord(article);

                Assert.Multiple(() =>
                {
                    Assert.IsNotNull(updatedFromQuery);
                    Assert.IsNotNull(updatedFromDatabase);
                    Assert.AreEqual(65, updatedFromDatabase.Ip);
                    Assert.AreEqual("2200", updatedFromDatabase.Height);
                    Assert.AreEqual("800", updatedFromDatabase.Width);
                    Assert.AreEqual("500", updatedFromDatabase.Depth);
                    Assert.AreEqual(productVendorId, updatedFromDatabase.ProductVendorId);
                    Assert.AreEqual("2200", updatedFromQuery.Height);
                    Assert.AreEqual(65, updatedFromQuery.Ip);
                });
            }
        }
    }
}

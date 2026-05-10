using ExcelMacroAdd.BusinessLayer.Models;
using NUnit.Framework;
using System;
using System.Threading.Tasks;

namespace ExelMacroAdd.Tests
{
    [TestFixture]
    public class NotPriceComponentsServiceIntegrationTests
    {
        [Test]
        public async Task AddRecordAsync_CreatesVendorAndPersistsRecord()
        {
            using (var scope = new SqliteIntegrationTestScope())
            {
                var service = scope.CreateNotPriceComponentsService();
                string article = "NP-" + TestContext.CurrentContext.Test.ID;
                string vendorName = "Vendor-" + Guid.NewGuid().ToString("N").Substring(0, 8);
                string multiplicityName = scope.GetFirstMultiplicityName();

                var created = await service.AddRecordAsync(new NotPriceComponentSaveRequest
                {
                    Article = article,
                    Description = "Integration test record",
                    ProductVendorName = vendorName,
                    MultiplicityName = multiplicityName,
                    Price = 123.45m,
                    Discount = 7,
                    Link = "https://example.test/item"
                }, createVendorIfMissing: true);

                var record = scope.GetNotPriceComponentRecord(article);
                var vendor = scope.GetProductVendor(vendorName);

                Assert.Multiple(() =>
                {
                    Assert.IsNotNull(created);
                    Assert.IsNotNull(record);
                    Assert.IsNotNull(vendor);
                    Assert.AreEqual(vendorName, record.ProductVendor.VendorName);
                    Assert.AreEqual(multiplicityName, record.Multiplicity.Value);
                    Assert.AreEqual(123.45m, record.Price);
                    Assert.AreEqual(7, record.Discount);
                });
            }
        }

        [Test]
        public async Task UpdateSetStateAndDelete_PersistRecordLifecycle()
        {
            using (var scope = new SqliteIntegrationTestScope())
            {
                var service = scope.CreateNotPriceComponentsService();
                string article = "NP-LIFE-" + TestContext.CurrentContext.Test.ID;
                string vendorName = "Vendor-" + Guid.NewGuid().ToString("N").Substring(0, 8);
                string multiplicityName = scope.GetFirstMultiplicityName();

                var created = await service.AddRecordAsync(new NotPriceComponentSaveRequest
                {
                    Article = article,
                    Description = "Original",
                    ProductVendorName = vendorName,
                    MultiplicityName = multiplicityName,
                    Price = 10m,
                    Discount = 1,
                    Link = "https://example.test/original"
                }, createVendorIfMissing: true);

                var updated = await service.UpdateRecordAsync(new NotPriceComponentSaveRequest
                {
                    Article = article,
                    Description = "Updated description",
                    ProductVendorName = vendorName,
                    MultiplicityName = multiplicityName,
                    Price = 20m,
                    Discount = 2,
                    Link = "https://example.test/updated"
                }, createVendorIfMissing: false);

                var stateChanged = await service.SetRecordStateAsync(article, 1);
                bool deleted = await service.DeleteRecordAsync(created.Id);
                var deletedRecord = scope.GetNotPriceComponentRecord(article);

                Assert.Multiple(() =>
                {
                    Assert.IsNotNull(updated);
                    Assert.AreEqual("Updated description", updated.Description);
                    Assert.AreEqual(20m, updated.Price);
                    Assert.AreEqual("https://example.test/updated", updated.Link);
                    Assert.IsNotNull(stateChanged);
                    Assert.AreEqual(1, stateChanged.IsValid);
                    Assert.IsTrue(deleted);
                    Assert.IsNull(deletedRecord);
                });
            }
        }
    }
}

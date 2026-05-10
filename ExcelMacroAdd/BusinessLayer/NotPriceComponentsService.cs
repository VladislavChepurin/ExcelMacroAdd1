using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.BusinessLayer.Models;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Threading.Tasks;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class NotPriceComponentsService : INotPriceComponentsService
    {
        private readonly IUnitOfWorkFactory _unitOfWorkFactory;

        public NotPriceComponentsService(IUnitOfWorkFactory unitOfWorkFactory)
        {
            _unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
        }

        public async Task<IList<NotPriceComponent>> GetAllRecordsAsync()
        {
            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                return await unitOfWork.Context.NotPriceComponents
                    .Include(p => p.ProductVendor)
                    .Include(p => p.Multiplicity)
                    .AsNoTracking()
                    .ToListAsync();
            }
        }

        public async Task<bool> RecordExistsAsync(string article)
        {
            if (string.IsNullOrWhiteSpace(article))
            {
                return false;
            }

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                return await unitOfWork.Context.NotPriceComponents.AnyAsync(p => p.Article == article);
            }
        }

        public async Task<bool> VendorExistsAsync(string vendorName)
        {
            if (string.IsNullOrWhiteSpace(vendorName))
            {
                return false;
            }

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                return await unitOfWork.Context.ProductVendors.AnyAsync(p => p.VendorName == vendorName);
            }
        }

        public async Task<NotPriceComponent> AddRecordAsync(NotPriceComponentSaveRequest request, bool createVendorIfMissing)
        {
            ValidateRequest(request);

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                var context = unitOfWork.Context;

                if (await context.NotPriceComponents.AnyAsync(p => p.Article == request.Article))
                {
                    throw new InvalidOperationException($"Артикул {request.Article} уже есть в базе данных.");
                }

                var productVendorId = await ResolveProductVendorIdAsync(unitOfWork, request.ProductVendorName, createVendorIfMissing);
                var multiplicityId = await ResolveMultiplicityIdAsync(context, request.MultiplicityName);

                var entity = new NotPriceComponent
                {
                    Article = request.Article,
                    Description = request.Description,
                    MultiplicityId = multiplicityId,
                    ProductVendorId = productVendorId,
                    Price = request.Price,
                    Discount = request.Discount,
                    DataRecord = CreateDataRecord(),
                    Link = request.Link
                };

                context.NotPriceComponents.Add(entity);
                await unitOfWork.SaveChangesAsync();
                return await ReloadRecordAsync(context, entity.Id);
            }
        }

        public async Task<NotPriceComponent> UpdateRecordAsync(NotPriceComponentSaveRequest request, bool createVendorIfMissing)
        {
            ValidateRequest(request);

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                var context = unitOfWork.Context;
                var existingRecord = await context.NotPriceComponents
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p => p.Article == request.Article);

                if (existingRecord == null)
                {
                    return null;
                }

                var updatedRecord = new NotPriceComponent
                {
                    Id = existingRecord.Id,
                    IsValid = existingRecord.IsValid,
                    Article = existingRecord.Article,
                    Description = request.Description,
                    MultiplicityId = await ResolveMultiplicityIdAsync(context, request.MultiplicityName),
                    ProductVendorId = await ResolveProductVendorIdAsync(unitOfWork, request.ProductVendorName, createVendorIfMissing),
                    Price = request.Price,
                    Discount = request.Discount,
                    DataRecord = CreateDataRecord(),
                    Link = string.IsNullOrWhiteSpace(request.Link) ? existingRecord.Link : request.Link
                };

                context.NotPriceComponents.Attach(updatedRecord);
                context.Entry(updatedRecord).State = EntityState.Modified;

                await unitOfWork.SaveChangesAsync();
                return await ReloadRecordAsync(context, updatedRecord.Id);
            }
        }

        public async Task<NotPriceComponent> SetRecordStateAsync(string article, int? status)
        {
            if (string.IsNullOrWhiteSpace(article))
            {
                return null;
            }

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                var entity = await unitOfWork.Context.NotPriceComponents
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p => p.Article == article);

                if (entity == null)
                {
                    return null;
                }

                entity.IsValid = status;
                unitOfWork.Context.NotPriceComponents.Attach(entity);
                unitOfWork.Context.Entry(entity).Property(p => p.IsValid).IsModified = true;
                await unitOfWork.SaveChangesAsync();
                return await ReloadRecordAsync(unitOfWork.Context, entity.Id);
            }
        }

        public async Task<bool> DeleteRecordAsync(int id)
        {
            try
            {
                using (var unitOfWork = _unitOfWorkFactory.Create())
                {
                    var entity = await unitOfWork.Context.NotPriceComponents
                        .FirstOrDefaultAsync(p => p.Id == id);

                    if (entity == null)
                    {
                        return false;
                    }

                    unitOfWork.Context.NotPriceComponents.Remove(entity);
                    await unitOfWork.SaveChangesAsync();
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static async Task<int?> ResolveProductVendorIdAsync(
            IUnitOfWork unitOfWork,
            string vendorName,
            bool createVendorIfMissing)
        {
            var context = unitOfWork.Context;
            var productVendor = await context.ProductVendors
                .FirstOrDefaultAsync(p => p.VendorName == vendorName);

            if (productVendor != null)
            {
                return productVendor.Id;
            }

            if (!createVendorIfMissing)
            {
                throw new InvalidOperationException($"В БД вендора '{vendorName}' нет.");
            }

            productVendor = new ProductVendor
            {
                VendorName = vendorName
            };

            context.ProductVendors.Add(productVendor);
            await unitOfWork.SaveChangesAsync();
            return productVendor.Id;
        }

        private static async Task<int?> ResolveMultiplicityIdAsync(AppContext context, string multiplicityName)
        {
            if (!string.IsNullOrWhiteSpace(multiplicityName))
            {
                var multiplicity = await context.Multiplicities
                    .FirstOrDefaultAsync(p => p.Value == multiplicityName);

                if (multiplicity != null)
                {
                    return multiplicity.Id;
                }
            }

            var defaultMultiplicity = await context.Multiplicities.FirstOrDefaultAsync(p => p.Id == 1);
            return defaultMultiplicity?.Id;
        }

        private static Task<NotPriceComponent> ReloadRecordAsync(AppContext context, int id)
        {
            return context.NotPriceComponents
                .Include(p => p.ProductVendor)
                .Include(p => p.Multiplicity)
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.Id == id);
        }

        private static string CreateDataRecord()
        {
            return DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");
        }

        private static void ValidateRequest(NotPriceComponentSaveRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.Article))
            {
                throw new ArgumentException("Артикул не может быть пустым.", nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.Description))
            {
                throw new ArgumentException("Описание не может быть пустым.", nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.ProductVendorName))
            {
                throw new ArgumentException("Вендор не может быть пустым.", nameof(request));
            }
        }
    }
}

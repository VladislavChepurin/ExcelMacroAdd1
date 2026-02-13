using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using Microsoft.Extensions.Caching.Memory;
using System.Collections.Generic;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessNotPriceComponent
    {
        private readonly AppContext context;
        private readonly IMemoryCache cache;

        public AccessNotPriceComponent(AppContext context, IMemoryCache cache)
        {
            this.context = context;
            this.cache = cache;
        }

        public async Task<IList<NotPriceComponent>> GetAllRecord()
        {           
            return await context.NotPriceComponents
                .Include(p => p.ProductVendor)
                .Include(p => p.Multiplicity)
                .AsNoTracking()
                .ToListAsync();
        }

        public async Task<bool> IsThereIsDBRecord (string аrticle)
        {   
            if (await context.NotPriceComponents.FirstOrDefaultAsync(p => p.Article == аrticle) is null)
            {
                return false;
            }
            return true;
        }

        public async Task AddValueDb(NotPriceComponent entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Detached;
                context.NotPriceComponents.Add(entity);
                await context.SaveChangesAsync();
            }
        }        
        
        public async Task<IProductVendor> AddProductVendor(ProductVendor vendor)
        {
            if (vendor != null)
            {
                context.ProductVendors.Add(vendor);
                await context.SaveChangesAsync();
                return vendor; // Возвращаем объект с обновленным ID
            }
            return null;
        }

        public async Task<IProductVendor> GetProductVendorEntityByName(string vendorName)
        {
            return await context.ProductVendors
                .FirstOrDefaultAsync(p => p.VendorName == vendorName);
        }

        public async Task<IMultiplicity> GetMultiplicityEntityByName(string multiplicityName)
        {
            return await context.Multiplicities
                .FirstOrDefaultAsync(p => p.Value == multiplicityName);
        }

        public async Task<bool> DeleteRecord(int id)
        {
            try
            {
                var entity = await context.NotPriceComponents
                    .FirstOrDefaultAsync(p => p.Id == id);

                if (entity == null)
                    return false;

                context.NotPriceComponents.Remove(entity);
                await context.SaveChangesAsync();
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        public async Task UpdateRecord(NotPriceComponent entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }
        public async Task<NotPriceComponent> GetRecordByArticle(string article)
        {
            return await context.NotPriceComponents
                .Include(p => p.ProductVendor)
                .FirstOrDefaultAsync(p => p.Article == article);
        }
    }
}

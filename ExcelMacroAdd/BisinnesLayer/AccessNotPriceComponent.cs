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
        }

        public async Task<IList<NotPriceComponent>> GetAllRecord()
        {           
            return await context.NotPriceComponents.Include(p => p.ProductVendor).AsNoTracking().ToListAsync();
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
                context.NotPriceComponents.Add(entity);
                await context.SaveChangesAsync();
            }
        }

        public async Task<IProductVendor> GetProductVendorEntityByName(string execution)
        {
            return await context.ProductVendors.FirstOrDefaultAsync(p => p.VendorName == execution);
        }

    }
}

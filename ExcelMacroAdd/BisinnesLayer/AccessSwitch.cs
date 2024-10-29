using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Models.Interface;
using Microsoft.Extensions.Caching.Memory;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessSwitch
    {
        private readonly AppContext context;
        private readonly IMemoryCache cache;
        public AccessSwitch(AppContext context, IMemoryCache cache)
        {
            this.context = context;
            this.cache = cache;
        }

        public async Task<ISwitch> GetEntitySwitch(string vendor, string series, int current, string quantityPole)
        {
            return await context.Switches
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.ProductVendor.VendorName == vendor
                                       && p.ProductSeries.SeriesName == series
                                       && p.Current == current
                                       && p.QuantityPole == quantityPole);
        }

        public string[] GetAllUniqueVendors()
        {
            return context.Switches
                .AsNoTracking()
                .Select(p => p.ProductVendor.VendorName)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetAllUniqueSeries(string vendor)
        {
            cache.TryGetValue(vendor, out string[] series);
            if (series == null)
            {
                series = context.Switches
                .AsNoTracking()
                .Where(p => p.ProductVendor.VendorName == vendor)
                .Select(p => p.ProductSeries.SeriesName)
                .ToHashSet()
                .ToArray();
                if (series != null)
                    cache.Set(vendor, series, new MemoryCacheEntryOptions().SetAbsoluteExpiration(System.TimeSpan.FromMinutes(5)));
            }
            return series;
        }

        public IUserSwitch GetDataSwitch(string vendor, string series)
        {
            var keyCache = string.Concat(vendor, series);

            cache.TryGetValue(keyCache, out IUserSwitch userSwitch);
            if (userSwitch == null)
            {
                var group = context.Switches
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.ProductGroup.GroupName)
                    .FirstOrDefault();

                var current = context.Switches
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.Current)
                    .OrderBy(p => p)
                    .ToHashSet()
                    .ToArray();

                var qantityPole = context.Switches
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.QuantityPole)
                    .ToHashSet()
                    .ToArray();

                userSwitch = new UserSwitch(group, current, qantityPole);
                if (userSwitch.current != null)
                    cache.Set(keyCache, userSwitch, new MemoryCacheEntryOptions().SetAbsoluteExpiration(System.TimeSpan.FromMinutes(5)));
            }
            return userSwitch;
        }
    }
}

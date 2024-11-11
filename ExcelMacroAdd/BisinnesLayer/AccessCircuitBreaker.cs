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
    public class AccessCircuitBreaker
    {
        private readonly AppContext context;
        private readonly IMemoryCache memoryCache;
        public AccessCircuitBreaker(AppContext context, IMemoryCache memoryCache)
        {
            this.context = context;
            this.memoryCache = memoryCache;
        }

        public async Task<ICircuitBreaker> GetEntityCircuitBreaker(string vendor, string series, int current, string curve, string maxCurrent, string quantityPole)
        {
            return await context.CircuitBreakers
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.ProductVendor.VendorName == vendor
                                       && p.ProductSeries.SeriesName == series
                                       && p.Current == current
                                       && p.Kurve == curve
                                       && p.MaxCurrent == maxCurrent
                                       && p.QuantityPole == quantityPole);
        }


        public string[] GetAllUniqueVendors()
        {
            return context.CircuitBreakers
                .AsNoTracking()
                .Select(p => p.ProductVendor.VendorName)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetAllUniqueSeries(string vendor)
        {
            var keyCache = string.Concat(vendor, "keyCircuitBreaker");

            memoryCache.TryGetValue(keyCache, out string[] series);
            if (series == null)
            {
                series = context.CircuitBreakers
               .AsNoTracking()
               .Where(p => p.ProductVendor.VendorName == vendor)
               .Select(p => p.ProductSeries.SeriesName)
               .ToHashSet()
               .ToArray();
                if (series != null)
                    memoryCache.Set(keyCache, series, new MemoryCacheEntryOptions().SetAbsoluteExpiration(System.TimeSpan.FromMinutes(5)));
            }
            return series;
        }

        public IUserCircuitBreaker GetDataCircutBreaker(string vendor, string series)
        {
            var keyCache = string.Concat(vendor, series, "keyCircuitBreaker");

            memoryCache.TryGetValue(keyCache, out IUserCircuitBreaker userCircuitBreaker);
            if (userCircuitBreaker == null)
            {
                var group = context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.ProductGroup.GroupName)
                    .FirstOrDefault();

                var current = context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.Current)
                    .OrderBy(p => p)
                    .ToHashSet()
                    .ToArray();

                var kurve = context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.Kurve)
                    .ToHashSet()
                    .ToArray();

                var maxCurrent = context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.MaxCurrent)
                    .ToHashSet()
                    .ToArray();

                var qantityPole = context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor && p.ProductSeries.SeriesName == series)
                    .Select(p => p.QuantityPole)
                    .ToHashSet()
                    .ToArray();

                userCircuitBreaker = new UserCircuitBreaker(group, current, kurve, maxCurrent, qantityPole);
                if (userCircuitBreaker.current != null)
                    memoryCache.Set(keyCache, userCircuitBreaker, new MemoryCacheEntryOptions().SetAbsoluteExpiration(System.TimeSpan.FromMinutes(5)));
            }
            return userCircuitBreaker;
        }
    }
}

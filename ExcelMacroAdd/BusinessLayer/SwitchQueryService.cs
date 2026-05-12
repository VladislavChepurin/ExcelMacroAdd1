using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Models.Interface;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class SwitchQueryService : ISwitchService
    {
        private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(10);
        private readonly IUnitOfWorkFactory unitOfWorkFactory;
        private readonly IMemoryCache cache;

        public SwitchQueryService(IUnitOfWorkFactory unitOfWorkFactory, IMemoryCache cache)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
            this.cache = cache ?? throw new ArgumentNullException(nameof(cache));
        }

        public async Task<ISwitch> GetEntitySwitch(string vendor, string series, int current, string quantityPole)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return await unitOfWork.Context.Switches
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p => p.ProductVendor.VendorName == vendor
                                           && p.ProductSeries.SeriesName == series
                                           && p.Current == current
                                           && p.QuantityPole == quantityPole);
            }
        }

        public string[] GetAllUniqueVendors()
        {
            return cache.GetOrCreate("switch:vendors", entry =>
            {
                entry.SetAbsoluteExpiration(CacheLifetime);

                using (var unitOfWork = unitOfWorkFactory.Create())
                {
                    return unitOfWork.Context.Switches
                        .AsNoTracking()
                        .Select(p => p.ProductVendor.VendorName)
                        .ToHashSet()
                        .ToArray();
                }
            });
        }

        public string[] GetAllUniqueSeries(string vendor)
        {
            string cacheKey = $"switch:series:{NormalizeCachePart(vendor)}";
            return cache.GetOrCreate(cacheKey, entry =>
            {
                entry.SetAbsoluteExpiration(CacheLifetime);

                using (var unitOfWork = unitOfWorkFactory.Create())
                {
                    return unitOfWork.Context.Switches
                        .AsNoTracking()
                        .Where(p => p.ProductVendor.VendorName == vendor)
                        .Select(p => p.ProductSeries.SeriesName)
                        .ToHashSet()
                        .ToArray();
                }
            });
        }

        public IUserSwitch GetDataSwitch(string vendor, string series)
        {
            string cacheKey = $"switch:data:{NormalizeCachePart(vendor)}:{NormalizeCachePart(series)}";
            return cache.GetOrCreate(cacheKey, entry =>
            {
                entry.SetAbsoluteExpiration(CacheLifetime);

                using (var unitOfWork = unitOfWorkFactory.Create())
                {
                    var items = unitOfWork.Context.Switches
                        .AsNoTracking()
                        .Where(p => p.ProductVendor.VendorName == vendor
                                 && p.ProductSeries.SeriesName == series)
                        .Select(p => new
                        {
                            Group = p.ProductGroup.GroupName,
                            p.Current,
                            p.QuantityPole
                        })
                        .ToList();

                    var group = items.Select(p => p.Group).FirstOrDefault();
                    var current = items.Select(p => p.Current).OrderBy(p => p).ToHashSet().ToArray();
                    var quantityPole = items.Select(p => p.QuantityPole).ToHashSet().ToArray();

                    return (IUserSwitch)new UserSwitch(group, current, quantityPole);
                }
            });
        }

        private static string NormalizeCachePart(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim().ToLowerInvariant();
        }
    }
}

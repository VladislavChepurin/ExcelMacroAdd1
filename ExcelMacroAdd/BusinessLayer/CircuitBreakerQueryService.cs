using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Models.Interface;
using System;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class CircuitBreakerQueryService : ICircuitBreakerService
    {
        private readonly IUnitOfWorkFactory unitOfWorkFactory;

        public CircuitBreakerQueryService(IUnitOfWorkFactory unitOfWorkFactory)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
        }

        public async Task<ICircuitBreaker> GetEntityCircuitBreaker(string vendor, string series, int current, string curve, string maxCurrent, string quantityPole)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return await unitOfWork.Context.CircuitBreakers
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p => p.ProductVendor.VendorName == vendor
                                           && p.ProductSeries.SeriesName == series
                                           && p.Current == current
                                           && p.Kurve == curve
                                           && p.MaxCurrent == maxCurrent
                                           && p.QuantityPole == quantityPole);
            }
        }

        public string[] GetAllUniqueVendors()
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.CircuitBreakers
                    .AsNoTracking()
                    .Select(p => p.ProductVendor.VendorName)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public string[] GetAllUniqueSeries(string vendor)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor)
                    .Select(p => p.ProductSeries.SeriesName)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public IUserCircuitBreaker GetDataCircutBreaker(string vendor, string series)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                var items = unitOfWork.Context.CircuitBreakers
                    .AsNoTracking()
                    .Where(p => p.ProductVendor.VendorName == vendor
                             && p.ProductSeries.SeriesName == series)
                    .Select(p => new
                    {
                        Group = p.ProductGroup.GroupName,
                        p.Current,
                        p.Kurve,
                        p.MaxCurrent,
                        p.QuantityPole
                    })
                    .ToList();

                var group = items.Select(p => p.Group).FirstOrDefault();
                var current = items.Select(p => p.Current).OrderBy(p => p).ToHashSet().ToArray();
                var kurve = items.Select(p => p.Kurve).ToHashSet().ToArray();
                var maxCurrent = items.Select(p => p.MaxCurrent).ToHashSet().ToArray();
                var quantityPole = items.Select(p => p.QuantityPole).ToHashSet().ToArray();

                return new UserCircuitBreaker(group, current, kurve, maxCurrent, quantityPole);
            }
        }
    }
}

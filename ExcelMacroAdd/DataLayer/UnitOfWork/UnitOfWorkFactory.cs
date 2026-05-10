using ExcelMacroAdd.DataLayer.Infrastructure;
using System;

namespace ExcelMacroAdd.DataLayer.UnitOfWork
{
    public sealed class UnitOfWorkFactory : IUnitOfWorkFactory
    {
        private readonly IAppContextFactory _appContextFactory;

        public UnitOfWorkFactory(IAppContextFactory appContextFactory)
        {
            _appContextFactory = appContextFactory ?? throw new ArgumentNullException(nameof(appContextFactory));
        }

        public IUnitOfWork Create()
        {
            return new EfUnitOfWork(_appContextFactory.Create());
        }
    }
}

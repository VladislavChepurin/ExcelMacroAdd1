using System;
using System.Threading;
using System.Threading.Tasks;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.DataLayer.UnitOfWork
{
    public sealed class EfUnitOfWork : IUnitOfWork
    {
        private readonly AppContext _context;
        private bool _disposed;

        public EfUnitOfWork(AppContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public AppContext Context => _context;

        public int SaveChanges()
        {
            return _context.SaveChanges();
        }

        public Task<int> SaveChangesAsync()
        {
            return _context.SaveChangesAsync();
        }

        public Task<int> SaveChangesAsync(CancellationToken cancellationToken)
        {
            return _context.SaveChangesAsync(cancellationToken);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _context.Dispose();
            _disposed = true;
        }
    }
}

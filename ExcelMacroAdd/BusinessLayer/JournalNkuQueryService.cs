using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class JournalNkuQueryService : IJournalNkuService
    {
        private static readonly object CacheMissMarker = new object();
        private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(5);

        private readonly IUnitOfWorkFactory unitOfWorkFactory;
        private readonly IMemoryCache cache;

        public JournalNkuQueryService(IUnitOfWorkFactory unitOfWorkFactory, IMemoryCache cache)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
            this.cache = cache ?? throw new ArgumentNullException(nameof(cache));
        }

        public async Task<IBoxBase> GetEntityJournal(string sArticle)
        {
            string normalizedArticle = NormalizeArticle(sArticle);
            if (string.IsNullOrEmpty(normalizedArticle))
            {
                return null;
            }

            if (TryGetCachedBoxBase(normalizedArticle, out IBoxBase boxBase))
            {
                return boxBase;
            }

            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                boxBase = await unitOfWork.Context.JornalNkus
                    .Include(p => p.MaterialBox)
                    .Include(p => p.ExecutionBox)
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p => p.Article == normalizedArticle) as IBoxBase;
            }

            SetCachedBoxBase(normalizedArticle, boxBase);
            return boxBase;
        }

        public async Task<IReadOnlyDictionary<string, IBoxBase>> GetEntityJournalBatch(IEnumerable<string> articles)
        {
            var normalizedArticles = articles?
                .Select(NormalizeArticle)
                .Where(article => !string.IsNullOrEmpty(article))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                ?? new List<string>();

            var result = new Dictionary<string, IBoxBase>(StringComparer.OrdinalIgnoreCase);
            var uncachedArticles = new List<string>();

            foreach (var article in normalizedArticles)
            {
                if (TryGetCachedBoxBase(article, out IBoxBase cachedBoxBase))
                {
                    result[article] = cachedBoxBase;
                }
                else
                {
                    uncachedArticles.Add(article);
                }
            }

            if (uncachedArticles.Count == 0)
            {
                return result;
            }

            List<BoxBase> entities;
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                entities = await unitOfWork.Context.JornalNkus
                    .Include(p => p.MaterialBox)
                    .Include(p => p.ExecutionBox)
                    .AsNoTracking()
                    .Where(p => uncachedArticles.Contains(p.Article))
                    .ToListAsync();
            }

            var entitiesByArticle = entities
                .Where(entity => !string.IsNullOrEmpty(entity.Article))
                .GroupBy(entity => NormalizeArticle(entity.Article))
                .ToDictionary(group => group.Key, group => (IBoxBase)group.First(), StringComparer.OrdinalIgnoreCase);

            foreach (var article in uncachedArticles)
            {
                entitiesByArticle.TryGetValue(article, out IBoxBase boxBase);
                SetCachedBoxBase(article, boxBase);
                result[article] = boxBase;
            }

            return result;
        }

        private bool TryGetCachedBoxBase(string article, out IBoxBase boxBase)
        {
            if (cache.TryGetValue(article, out object cachedValue))
            {
                boxBase = ReferenceEquals(cachedValue, CacheMissMarker)
                    ? null
                    : cachedValue as IBoxBase;
                return true;
            }

            boxBase = null;
            return false;
        }

        private void SetCachedBoxBase(string article, IBoxBase boxBase)
        {
            cache.Set(
                article,
                boxBase ?? CacheMissMarker,
                new MemoryCacheEntryOptions().SetAbsoluteExpiration(CacheLifetime));
        }

        private static string NormalizeArticle(string article)
        {
            return string.IsNullOrWhiteSpace(article)
                ? null
                : article.Trim().ToLowerInvariant();
        }
    }
}

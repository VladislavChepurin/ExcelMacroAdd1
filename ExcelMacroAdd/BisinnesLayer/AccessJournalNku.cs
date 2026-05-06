using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessJournalNku
    {
        private static readonly object CacheMissMarker = new object();
        private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(5);

        private readonly AppContext context;
        private readonly IMemoryCache cache;

        public AccessJournalNku(AppContext context, IMemoryCache cache)
        {
            this.context = context;
            this.cache = cache;
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

            boxBase = await context.JornalNkus
                .Include(p => p.MaterialBox)
                .Include(p => p.ExecutionBox)
                .FirstOrDefaultAsync(p => p.Article == normalizedArticle) as IBoxBase;

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

            var entities = await context.JornalNkus
                .Include(p => p.MaterialBox)
                .Include(p => p.ExecutionBox)
                .Where(p => uncachedArticles.Contains(p.Article))
                .ToListAsync();

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

        public async Task WriteUpdateDb(BoxBase entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }

        public async Task<IMaterialBox> GetMaterialEntityByName(string material)
        {
            return await context.Materials.FirstOrDefaultAsync(p => p.MaterialValue == material);
        }

        public async Task<IExecutionBox> GetExecutionEntityByName(string execution)
        {
            return await context.Executions.FirstOrDefaultAsync(p => p.ExecutionValue == execution);
        }

        public async Task AddValueDb(BoxBase entity)
        {
            if (entity != null)
            {
                context.JornalNkus.Add(entity);
                await context.SaveChangesAsync();
            }
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

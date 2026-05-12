using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.BusinessLayer.Models;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using ExcelMacroAdd.UserException;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Data.Entity;
using System.Threading.Tasks;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class JournalNkuWriteService : IJournalNkuWriteService
    {
        private readonly IUnitOfWorkFactory _unitOfWorkFactory;
        private readonly IMemoryCache _memoryCache;

        public JournalNkuWriteService(IUnitOfWorkFactory unitOfWorkFactory, IMemoryCache memoryCache)
        {
            _unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
            _memoryCache = memoryCache ?? throw new ArgumentNullException(nameof(memoryCache));
        }

        public async Task<JournalNkuWriteResult> AddBoxAsync(JournalNkuWriteRequest request)
        {
            ValidateRequest(request);

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                var context = unitOfWork.Context;
                string article = NormalizeArticle(request.Article);

                bool exists = await context.JornalNkus
                    .AsNoTracking()
                    .AnyAsync(p => p.Article == article);

                if (exists)
                {
                    return new JournalNkuWriteResult(JournalNkuWriteStatus.AlreadyExists, article);
                }

                var materialId = await ResolveMaterialIdAsync(context, request.MaterialName);
                var executionId = await ResolveExecutionIdAsync(context, request.ExecutionName);

                var entity = new BoxBase
                {
                    Ip = request.Ip,
                    Climate = NormalizeOptionalValue(request.Climate),
                    Weight = NormalizeOptionalValue(request.Weight),
                    Height = request.Height,
                    Width = request.Width,
                    Depth = request.Depth,
                    Article = article,
                    MaterialBoxId = materialId,
                    ExecutionBoxId = executionId
                };

                context.JornalNkus.Add(entity);
                await unitOfWork.SaveChangesAsync();
                InvalidateCache(article);
                return new JournalNkuWriteResult(JournalNkuWriteStatus.Added, article);
            }
        }

        public async Task<JournalNkuWriteResult> UpdateBoxAsync(JournalNkuWriteRequest request)
        {
            ValidateRequest(request);

            using (var unitOfWork = _unitOfWorkFactory.Create())
            {
                var context = unitOfWork.Context;
                string article = NormalizeArticle(request.Article);

                var existing = await context.JornalNkus
                    .FirstOrDefaultAsync(p => p.Article == article);

                if (existing == null)
                {
                    return new JournalNkuWriteResult(JournalNkuWriteStatus.NotFound, article);
                }

                var materialId = await ResolveMaterialIdAsync(context, request.MaterialName);
                var executionId = await ResolveExecutionIdAsync(context, request.ExecutionName);

                existing.Article = article;
                existing.Ip = request.Ip;
                existing.Climate = NormalizeOptionalValue(request.Climate);
                existing.Weight = NormalizeOptionalValue(request.Weight);
                existing.Height = request.Height;
                existing.Width = request.Width;
                existing.Depth = request.Depth;
                existing.MaterialBoxId = materialId;
                existing.ExecutionBoxId = executionId;

                await unitOfWork.SaveChangesAsync();
                InvalidateCache(article);
                return new JournalNkuWriteResult(JournalNkuWriteStatus.Updated, article);
            }
        }

        private static async Task<int?> ResolveMaterialIdAsync(AppContext context, string materialName)
        {
            var material = await context.Materials
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.MaterialValue == materialName);

            if (material == null)
            {
                throw new DataBaseNotFoundValueException(
                    $"Введенный материал шкафа \"{materialName}\" недопустим, пожалуйста используйте значение \"Пластик\", или  \"Металл\", или \"Композит\"");
            }

            return material.Id;
        }

        private static async Task<int?> ResolveExecutionIdAsync(AppContext context, string executionName)
        {
            var execution = await context.Executions
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.ExecutionValue == executionName);

            if (execution == null)
            {
                throw new DataBaseNotFoundValueException(
                    $"Введенное исполнение шкафа \"{executionName}\" недопустимо, пожалуйста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");
            }

            return execution.Id;
        }

        private void InvalidateCache(string article)
        {
            if (!string.IsNullOrEmpty(article))
            {
                _memoryCache.Remove(article);
            }
        }

        private static void ValidateRequest(JournalNkuWriteRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.Article))
            {
                throw new ArgumentException("Артикул не может быть пустым.", nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.Height)
                || string.IsNullOrWhiteSpace(request.Width)
                || string.IsNullOrWhiteSpace(request.Depth)
                || string.IsNullOrWhiteSpace(request.MaterialName)
                || string.IsNullOrWhiteSpace(request.ExecutionName))
            {
                throw new ArgumentException("Не заполнены обязательные поля записи.", nameof(request));
            }
        }

        private static string NormalizeOptionalValue(string value)
        {
            return value == "-" ? null : value;
        }

        private static string NormalizeArticle(string article)
        {
            return string.IsNullOrWhiteSpace(article)
                ? null
                : article.Trim().ToLowerInvariant();
        }
    }
}

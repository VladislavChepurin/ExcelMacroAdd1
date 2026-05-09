using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace ExcelMacroAdd.Functions
{
    /// <summary>
    /// Управляет производительностью Excel при пакетных операциях.
    /// 
    /// Два режима использования:
    /// 
    /// 1) Как IDisposable-обёртка (using-блок):
    ///    using (var scope = new ExcelPerformanceScope(Application))
    ///    {
    ///        // Здесь ScreenUpdating=false, Calculation=Manual, EnableEvents=false
    ///        // Можно вставлять формулы без пересчёта
    ///    }
    ///    // При Dispose: восстановление настроек + однократный Calculate
    ///    //              (только если лист ещё не пересчитывался в этой сессии)
    ///
    /// 2) Вложенное использование (формы вызывают WriteExcel.Start() в цикле):
    ///    Внутренний scope видит, что внешний уже активен, и НЕ восстанавливает
    ///    настройки при своём Dispose — это сделает внешний scope.
    ///    
    /// Кэш пересчитанных листов живёт на уровне Application (статический),
    /// сбрасывается при смене книги или по таймауту.
    /// </summary>
    public sealed class ExcelPerformanceScope : IDisposable
    {
        private readonly Application _app;
        private readonly bool _isOuterScope;
        private readonly bool _originalScreenUpdating;
        private readonly XlCalculation _originalCalculation;
        private readonly bool _originalEnableEvents;
        private bool _disposed;
        private static readonly object _bookChangeLock = new object();

        // ===============================================================
        // Статический кэш: какие листы уже пересчитывались в этой сессии.
        // Ключ = "BookName|SheetName", чтобы различать листы разных книг.
        // ===============================================================
        private static readonly ConcurrentDictionary<string, byte> _calculatedSheets
            = new ConcurrentDictionary<string, byte>();
        private static volatile string _lastWorkbookFullName;

        // Счётчик вложенности — позволяет определить, кто "внешний"
        [ThreadStatic]
        private static int _nestingLevel;

        public static int CurrentNestingLevel => _nestingLevel;

        public ExcelPerformanceScope(Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            _nestingLevel++;
            _isOuterScope = (_nestingLevel == 1);

            // Сохраняем оригинальные настройки только на внешнем уровне
            if (_isOuterScope)
            {
                _originalScreenUpdating = _app.ScreenUpdating;
                _originalCalculation = _app.Calculation;
                _originalEnableEvents = _app.EnableEvents;

                _app.ScreenUpdating = false;
                _app.Calculation = XlCalculation.xlCalculationManual;
                _app.EnableEvents = false;
            }

            // Проверяем, не сменилась ли книга — если да, сбрасываем кэш
            InvalidateCacheIfBookChanged();
        }

        /// <summary>
        /// Был ли текущий лист уже пересчитан в этой сессии?
        /// </summary>
        public bool IsCurrentSheetAlreadyCalculated
        {
            get
            {
                string key = GetCurrentSheetKey();
                return key != null && _calculatedSheets.ContainsKey(key);
            }
        }

        /// <summary>
        /// Принудительно пересчитать текущий лист и пометить его как пересчитанный.
        /// Вызывается вручную, если нужно гарантировать актуальность данных.
        /// </summary>
        public void ForceCalculateCurrentSheet()
        {
            try
            {
                var sheet = _app.ActiveSheet as Worksheet;
                sheet?.Calculate();
                MarkCurrentSheetCalculated();
            }
            catch { /* Игнорируем ошибки пересчёта */ }
        }

        /// <summary>
        /// Пометить текущий лист как пересчитанный (без фактического пересчёта).
        /// </summary>
        public void MarkCurrentSheetCalculated()
        {
            string key = GetCurrentSheetKey();
            if (key != null)
                _calculatedSheets.TryAdd(key, 0);
        }

        /// <summary>
        /// Сбросить кэш пересчитанных листов (например, при загрузке нового прайса).
        /// </summary>
        public static void InvalidateCache()
        {
            _calculatedSheets.Clear();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _nestingLevel--;

            // Восстанавливаем настройки только на внешнем уровне
            if (_isOuterScope)
            {
                // Пересчитываем только если:
                // 1) Исходно был автопересчёт
                // 2) Лист ещё не пересчитывался в этой сессии
                if (_originalCalculation == XlCalculation.xlCalculationAutomatic)
                {
                    if (!IsCurrentSheetAlreadyCalculated)
                    {
                        try { _app.Calculate(); } catch { }
                        MarkCurrentSheetCalculated();
                    }
                }

                _app.EnableEvents = _originalEnableEvents;
                _app.Calculation = _originalCalculation;
                _app.ScreenUpdating = _originalScreenUpdating;
            }
        }

        // ===============================================================
        // Приватные вспомогательные методы
        // ===============================================================

        private string GetCurrentSheetKey()
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                var ws = _app.ActiveSheet as Worksheet;
                if (wb == null || ws == null) return null;
                return $"{wb.FullName}|{ws.Name}";
            }
            catch
            {
                return null;
            }
        }

        private void InvalidateCacheIfBookChanged()
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return;

                string currentBook = wb.FullName;
                lock (_bookChangeLock)
                {
                    if (_lastWorkbookFullName != currentBook)
                    {
                        _calculatedSheets.Clear();
                        _lastWorkbookFullName = currentBook;
                    }
                }
            }
            catch { }
        }
    }
}
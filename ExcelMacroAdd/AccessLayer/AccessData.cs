using System;
using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;
using System.Collections.Generic;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData : IForm2Data, IJournalData, IForm4Data
    {
        public async Task<ISwitch> GetEntitySwitch(string current, string quantity)
        {
            using (var db = new AppContext())
            {
                return await db.Switchs.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Quantity == quantity);
            }
        }

        public async Task<IModul> GetEntityModule(string current, string curve, string maxCurrent, string quantity)
        {
            using (var db = new AppContext())
            {
                return await db.Moduls.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Kurve == curve && p.MaxCurrent == maxCurrent && p.Quantity == quantity);
            }
        }

        public async Task<IJournalNku> GetEntityJournal(string sArticle)
        {
            using (var db = new AppContext())
            {
                return await db.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
            }
        }

        public async void WriteUpdateDB(JournalNku entity)
        {
            using (var db = new AppContext())
            {
                if (entity != null)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    await db.SaveChangesAsync();
                }
            }
        }

        public async void AddValueDB(JournalNku entity)
        {
            using (var db = new AppContext())
            {
                if (entity != null)
                {    
                    db.JornalNkus.Add(entity);
                    await db.SaveChangesAsync();
                }
            }
        }

        public string[] GetComboBox2Items(string current)
        {
            using (var db = new AppContext())
            {
                return db.Transformers
                    .Where(p => p.Current == current)
                    .Select(p => p.Bus)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public string[] GetComboBox3Items(string current, string bus)
        {
            using (var db = new AppContext())
            {
                return db.Transformers
                    .Where(p => p.Current == current && p.Bus == bus)
                    .Select(p => p.Accuracy)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public string[] GetComboBox4Items(string current, string bus, string accuracy)
        {
            using (var db = new AppContext())
            {
                return db.Transformers
                    .Where(p => p.Current == current && p.Bus == bus && p.Accuracy == accuracy)
                    .Select(p => p.Power)
                    .ToHashSet()
                    .ToArray();
            }
        }
    }
}

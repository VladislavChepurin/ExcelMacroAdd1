using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using ExcelMacroAdd.UserVariables;
using System;
using System.Collections.Concurrent;

//Rewiew OK 22.04.2025
namespace ExcelMacroAdd.ProxyObjects
{    internal class DataInXmlProxy : IDataInXml
    {
        private readonly IDataInXml _dataInXml;
        private readonly ConcurrentDictionary<string, Vendor> _cache = new ConcurrentDictionary<string, Vendor>();
        private Vendor[] vendors;      

        public DataInXmlProxy(DataInXml dataInXml)
        {
            this._dataInXml = dataInXml;
        }

        public Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue)
        {
            if (vendor == null)
            {
                throw new ArgumentNullException(nameof(vendor));
            }

            string cacheKey = vendor.ToUpperInvariant();
            return _cache.GetOrAdd(cacheKey, _ => _dataInXml.ReadElementXml(vendor, ReadFileXml()));
        }

        public Vendor[] ReadFileXml()
        {
            if (vendors == null)
            {
                vendors = _dataInXml.ReadFileXml();
                return vendors;
            }
            return vendors;
        }

        public void WriteXml(string vendor, params string[] data)
        {
            //Очищаем 
            _cache.Clear();
            vendors = null;
            //Проксируем вызов на прямую
            _dataInXml.WriteXml(vendor, data);
        }

        public void XmlFileCreate()
        {
            //Проксируем вызов на прямую
            _dataInXml.XmlFileCreate();
        }
    }
}

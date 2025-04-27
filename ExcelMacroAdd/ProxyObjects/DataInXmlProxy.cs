using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using ExcelMacroAdd.UserVariables;
using System.Collections.Concurrent;
using System.Linq;

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
            return _cache.GetOrAdd(vendor, key =>
            {
                var vendors = _dataInXml.ReadFileXml();
                return vendors.Single(p => p.VendorAttribute == key);
            });         
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

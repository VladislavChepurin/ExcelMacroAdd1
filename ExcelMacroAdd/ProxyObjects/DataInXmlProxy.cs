using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using ExcelMacroAdd.UserVariables;
using System;
using System.Collections.Generic;

namespace ExcelMacroAdd.ProxyObjects
{

    internal class DataInXmlProxy : IDataInXml
    {
        private readonly Lazy<DataInXml> _dataXml;
        private readonly IDictionary<string, Vendor> _cacheSeveralXmlrecords = new Dictionary<string, Vendor>();

        public DataInXmlProxy(Lazy<DataInXml> dataXml)
        {
            _dataXml = dataXml;
        }

        public Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue)
        {
            if (!_cacheSeveralXmlrecords.ContainsKey(vendor))
            {
                var value = _dataXml.Value.ReadElementXml(vendor, _dataXml.Value.ReadFileXml());
                _cacheSeveralXmlrecords.Add(vendor, (Vendor)value);
                return value;
            }
            return _cacheSeveralXmlrecords[vendor];
        }

        public Vendor[] ReadFileXml()
        {
            //Проксируем вызов на прямую
            return _dataXml.Value.ReadFileXml();
        }

        public void WriteXml(string vendor, params string[] data)
        {
            //Очищаем коллекцию
            _cacheSeveralXmlrecords.Clear();
            //Проксируем вызов на прямую
            _dataXml.Value.WriteXml(vendor, data);
        }

        public void XmlFileCreate()
        {
            //Проксируем вызов на прямую
            _dataXml.Value.XmlFileCreate();
        }
    }
}

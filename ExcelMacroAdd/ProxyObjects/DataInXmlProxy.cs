using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.UserVariables;
using System;
using System.Collections.Generic;
using ExcelMacroAdd.Services;

namespace ExcelMacroAdd.ProxyObjects
{

    internal class DataInXmlProxy : IDataInXml
    {
        private readonly Lazy<DataInXml> dataXml;
        private readonly IDictionary<string, Vendor> cacheSeveralXmlRecords = new Dictionary<string, Vendor>();

        public DataInXmlProxy(Lazy<DataInXml> dataXml)
        {
            this.dataXml = dataXml;
        }

        public Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue)
        {
            if (!cacheSeveralXmlRecords.ContainsKey(vendor))
            {
                var value = dataXml.Value.ReadElementXml(vendor, dataXml.Value.ReadFileXml());
                cacheSeveralXmlRecords.Add(vendor, value);
                return value;
            }
            return cacheSeveralXmlRecords[vendor];
        }

        public Vendor[] ReadFileXml()
        {
            //Проксируем вызов на прямую
            return dataXml.Value.ReadFileXml();
        }

        public void WriteXml(string vendor, params string[] data)
        {
            //Очищаем коллекцию
            cacheSeveralXmlRecords.Clear();
            //Проксируем вызов на прямую
            dataXml.Value.WriteXml(vendor, data);
        }

        public void XmlFileCreate()
        {
            //Проксируем вызов на прямую
            dataXml.Value.XmlFileCreate();
        }
    }
}

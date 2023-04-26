using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Xml.Serialization;
using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Services
{
    public class DataInXml: IDataInXml
    {
        // Folders AppData content Settings.xml
        readonly string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/Settings.xml");
        public Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue)
        {
            return dataXmlContinue.Single(p => p.VendorAttribute == ReplaceVendorTable()[vendor]);
        }

        public Vendor[] ReadFileXml()
        {      
            XmlAttributes attrs = new XmlAttributes();
            XmlAttributeOverrides xOver = new XmlAttributeOverrides();
            XmlRootAttribute xRoot = new XmlRootAttribute
            {
                // Set a new Namespace and ElementName for the root element.
                ElementName = "MetaSettings"
            };
            attrs.XmlRoot = xRoot;
            xOver.Add(typeof(Vendor[]), attrs);
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Vendor[]), xOver);

            try 
            {
                // десериализуем
                using (FileStream fs = new FileStream(file, FileMode.OpenOrCreate))
                {
                    return xmlSerializer.Deserialize(fs) as Vendor[];                  
                }
            }
            catch (InvalidOperationException)
            {
                XmlFileCreate();
                return default;
            }
        }

        public void WriteXml(string vendor, params string[] data)
        {
            XDocument xdoc = XDocument.Load(file);
            var index = xdoc.Element("MetaSettings")?.Elements("Vendor").FirstOrDefault(p => p.Attribute("vendor")?.Value == vendor);
            if (index != null)
            {
                // Записываем первую формулу
                var formula1 = index.Element("Formula_1");
                if (formula1 != null) formula1.Value = data[0];
                // Записываем вторую формулу
                var formula2 = index.Element("Formula_2");
                if (formula2 != null) formula2.Value = data[1];
                // Записываем третью формулу
                var formula3 = index.Element("Formula_3");
                if (formula3 != null) formula3.Value = data[2];
                // Записываем скидку
                var discount = index.Element("Discount");
                if (discount != null) discount.Value = data[3];





                // Записываем дату и время
                var date = index.Element("Date");
                if (date != null) date.Value = data[4];        
                // Сохраняем документ
                xdoc.Save(file);
            }
        }

        public void XmlFileCreate()
        {    
            XmlAttributes attrs = new XmlAttributes();
            XmlAttributeOverrides xOver = new XmlAttributeOverrides();
            XmlRootAttribute xRoot = new XmlRootAttribute
            {
                // Set a new Namespace and ElementName for the root element.
                ElementName = "MetaSettings"
            };
            attrs.XmlRoot = xRoot;
            xOver.Add(typeof(Vendor[]), attrs);           

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Vendor[]), xOver);

            Vendor[] vendor = {
                new Vendor("IEK", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("EKF", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("DKC", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("KEAZ", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("DEKraft", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("TDM", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("ABB", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("Schneider", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4),
                new Vendor("Chint", "_", "_", "_", 0, DateTime.Now.ToString(new CultureInfo("ru-RU")), 1, 2, 3, 4)
            };
            // получаем поток, куда будем записывать сериализованный объект
            using (FileStream fs = new FileStream(file, FileMode.OpenOrCreate))
            {
                xmlSerializer.Serialize(fs, vendor);    
            }
        }

        /// <summary>
        /// Функция замены для вставки вендора и запроса из XML
        /// </summary>
        /// <returns></returns>
        public IDictionary<string, string> ReplaceVendorTable()                         
        {
            Dictionary<string, string> dictionaryVendor = new Dictionary<string, string>()
            {
                {"Iek", "IEK"},
                {"Ekf", "EKF"},
                {"IekVa47", "IEK"},
                {"IekVa47m", "IEK"},
                {"IekArmat", "IEK"},
                {"EkfProxima", "EKF"},
                {"EkfAvers", "EKF"},
                {"Abb", "ABB"},
                {"Keaz", "KEAZ"},
                {"Dkc", "DKC"},
                {"Dekraft", "DEKraft"},
                {"Schneider", "Schneider"},
                {"Chint", "Chint"},
                {"Tdm", "TDM"}               
            };
            return dictionaryVendor;
        }
    }
}

using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.UserVariables;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ExcelMacroAdd.Servises
{
    public class DataInXml: IDataInXml
    {
        // Folders AppData content Settings.xml
        readonly string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns\ExcelMacroAdd\Settings.xml");
        public Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue)
        {
            return dataXmlContinue.Where(p => p.VendorAttribute == Replace.RepleceVendorTable(vendor)).Single();
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
                // десериализуем объект
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
                var formula_1 = index.Element("Formula_1");
                if (formula_1 != null) formula_1.Value = data[0];
                // Записываем вторую формулу
                var formula_2 = index.Element("Formula_2");
                if (formula_2 != null) formula_2.Value = data[1];
                // Записываем третью формулу
                var formula_3 = index.Element("Formula_3");
                if (formula_3 != null) formula_3.Value = data[2];
                // Записываем скидку
                var discont = index.Element("Discont");
                if (discont != null) discont.Value = data[3];
                // Записываем дату и время
                DateTime localDate = DateTime.Now;
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

            Vendor[] vendor = new Vendor[8]
            {
                new Vendor("IEK", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("EKF", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("DKC", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("KEAZ", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("DEKraft", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("TDM", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("ABB", "_", "_", "_", 0, DateTime.Now.ToString()),
                new Vendor("Schneider", "_", "_", "_", 0, DateTime.Now.ToString())
            };
            // получаем поток, куда будем записывать сериализованный объект
            using (FileStream fs = new FileStream(file, FileMode.OpenOrCreate))
            {
                xmlSerializer.Serialize(fs, vendor);    
            }
        }
    }
}

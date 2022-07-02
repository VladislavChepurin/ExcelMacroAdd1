using ExcelMacroAdd.UserVariables;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ExcelMacroAdd.Servises
{
    public class DataInXml
    {
        // Folders AppData content Settings.xml
        readonly string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns\ExcelMacroAdd\Settings.xml");
        public string Vendor { get; set; }  
             
        public string ReadFileXml(string element)
        {   
            string middle = default;
            try
            {    
                XDocument xdoc = XDocument.Load(file);
                var toDiscont = xdoc.Element("MetaSettings")?   // получаем корневой узел MetaSettings
                    .Elements("Vendor")                         // получаем все элементы Vendor                               
                    .Where(p => p.Attribute("vendor")?.Value == Replace.RepleceVendorTable(Vendor))
                    .Select(p => new                            // для каждого объекта создаем анонимный объект
                    {
                        dataXml = p.Element(element)?.Value
                    });

                if (toDiscont != null)
                {
                    foreach (var data in toDiscont)
                    {
                        middle = data.dataXml;
                    }
                }
                return middle ?? String.Empty;
            }
            catch (Exception)
            {
                return String.Empty;
            }
        }

        public Vendor[] ReadFileXmlNew()
        {
            //  throw new Exception("Метод не реализован");

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

            // десериализуем объект
            using (FileStream fs = new FileStream("person.xml", FileMode.OpenOrCreate))
            {
              return  xmlSerializer.Deserialize(fs) as Vendor[];                
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

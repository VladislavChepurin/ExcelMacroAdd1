using System;
using System.Linq;
using System.Xml.Linq;

namespace ExcelMacroAdd.Servises
{
    internal class DataInXml
    {
        // Folders AppData content Settings.xml
        readonly string file = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\AddIns\ExcelMacroAdd\Settings.xml";
        public string Vendor { get; set; }  
             
        public string GetDataInXml(string element)
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

        public void XmlFileCreate()
        {
            XDocument xdoc = new XDocument(new XElement("MetaSettings",
                 //Поле IEK
                 new XElement("Vendor",
                 new XAttribute("vendor", "IEK"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле EKF
                 new XElement("Vendor",
                 new XAttribute("vendor", "EKF"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле DKC
                 new XElement("Vendor",
                 new XAttribute("vendor", "DKC"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле KEAZ
                 new XElement("Vendor",
                 new XAttribute("vendor", "KEAZ"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле DEKraft
                 new XElement("Vendor",
                 new XAttribute("vendor", "DEKraft"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле TDM
                 new XElement("Vendor",
                 new XAttribute("vendor", "TDM"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле ABB
                 new XElement("Vendor",
                 new XAttribute("vendor", "ABB"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_")),
                 //Поле Schneider
                 new XElement("Vendor",
                 new XAttribute("vendor", "Schneider"),
                 new XElement("Formula_1", "_"),
                 new XElement("Formula_2", "_"),
                 new XElement("Formula_3", "_"),
                 new XElement("Discont", "_"),
                 new XElement("Date", "_"))));
            xdoc.Save(file);
        }
    }
}

using System;
using System.Xml.Serialization;

namespace ExcelMacroAdd.UserVariables
{
    [Serializable]
    public class Vendor
    {
        [XmlAttribute("vendor")]
        public string VendorAttribute { get; set; }

        [XmlElement("Formula_1")]
        public string Formula_1 { get; set; }

        [XmlElement("Formula_2")]
        public string Formula_2 { get; set; }

        [XmlElement("Formula_3")]
        public string Formula_3 { get; set; }

        [XmlElement("Discount")]
        public int Discount { get; set; }

        [XmlElement("Date")]
        public string Date { get; set; }

        public Vendor() {  }

        public Vendor(string vendorAttribute, string formula_1, string formula_2, string formula_3, int discount, string date)
        {
            VendorAttribute = vendorAttribute;
            Formula_1 = formula_1;
            Formula_2 = formula_2;
            Formula_3 = formula_3;
            Discount = discount;
            Date = date;
        }
    }
}

using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Interfaces
{
    interface IDataInXml
    {
        Vendor ReadElementXml(string vendor);
        Vendor[] ReadFileXml();
        void WriteXml(string vendor, params string[] data);
        void XmlFileCreate();
    }
}

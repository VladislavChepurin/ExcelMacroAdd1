using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Interfaces
{
    public interface IDataInXml
    {
        ICurrentVendor ReadElementXml(string vendor, Vendor[] dataXmlContinue);
        Vendor[] ReadFileXml();
        void WriteXml(string vendor, params string[] data);
        void XmlFileCreate();
    }
}

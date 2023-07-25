using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Services.Interfaces
{
    public interface IDataInXml
    {
        Vendor ReadElementXml(string vendor, Vendor[] dataXmlContinue);
        Vendor[] ReadFileXml();
        void WriteXml(string vendor, params string[] data);
        void XmlFileCreate();
    }
}

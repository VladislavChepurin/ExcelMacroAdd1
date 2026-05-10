using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface ITransformerService
    {
        string[] GetComboBox2Items(string current);

        string[] GetComboBox3Items(string current, string bus);

        string[] GetComboBox4Items(string current, string bus, string accuracy);

        StructTransformer GetArticle(string current, string bus, string accuracy, string power);

        byte[] GetBlobPictureDb(string current, string bus, string accuracy, string power);
    }
}

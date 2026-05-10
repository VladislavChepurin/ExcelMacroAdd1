namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface ITwinBlockService
    {
        string[] GetComboBox1Items();

        (string, string, string, string, string) GetDataInTableDb(string current, bool isReverse);

        byte[] GetBlobPictureDb(string current, bool isReverse);
    }
}

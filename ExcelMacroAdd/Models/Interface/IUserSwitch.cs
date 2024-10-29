namespace ExcelMacroAdd.Models.Interface
{
    public interface IUserSwitch
    {

        string group { get; set; }
        int[] current { get; set; }
        string[] quantityPole { get; set; }

    }
}

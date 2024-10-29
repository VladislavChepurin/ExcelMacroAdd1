namespace ExcelMacroAdd.Models.Interface
{
    public interface IUserCircuitBreaker
    {

        string group { get; set; }
        int[] current { get; set; }
        string[] kurve { get; set; }
        string[] maxCurrent { get; set; }
        string[] quantityPole { get; set; }

    }
}

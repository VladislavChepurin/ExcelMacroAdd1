using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class Execution : IExecution
    {
        public int Id { get; set; }
        public string ExecutionValue { get; set; }
    }
}

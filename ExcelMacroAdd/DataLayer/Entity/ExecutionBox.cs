using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class ExecutionBox : IExecutionBox
    {
        public int Id { get; set; }
        public string ExecutionValue { get; set; }
    }
}

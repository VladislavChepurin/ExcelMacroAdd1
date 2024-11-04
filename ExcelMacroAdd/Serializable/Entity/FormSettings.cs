using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class FormSettings : IFormSettings
    {
        public bool FormTopMost { get; set; }

        public FormSettings(bool formTopMost)
        {
            FormTopMost = formTopMost;
        }
    }
}

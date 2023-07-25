using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class TypeNkySettings : ITypeNkySettings
    {
        public int Number { get; set; }
        public string Description { get; set; }
        public float BuildTime { get; set; }

        public TypeNkySettings(int number, string description, float buildTime)
        {
            Number = number;
            Description = description;
            BuildTime = buildTime;
        }
    }
}

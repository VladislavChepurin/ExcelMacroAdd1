using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class ResourcesForm4: IResourcesForm4
    {
        public object[] TransformerCurrent { get; set; }

        public ResourcesForm4(object[] transformerCurrent)
        {
            TransformerCurrent = transformerCurrent;
        }
    }
}

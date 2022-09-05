using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable
{
    public class ResourcesForm1 : IResourcesForm1
    {
        public int HeihgtMaxBox { get; set; }
        public string TempleteWall { get; set; }
        public string TempleteFloor { get; set; }

        public ResourcesForm1(int heihgtMaxBox, string templeteWall, string templeteFloor)
        {
            HeihgtMaxBox = heihgtMaxBox;
            TempleteWall = templeteWall;
            TempleteFloor = templeteFloor;
        }
    }
}

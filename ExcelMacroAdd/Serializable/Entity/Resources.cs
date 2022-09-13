using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class Resources : IResources
    {
        public string NameFileJornal { get; set; }
        public int HeihgtMaxBox { get; set; }
        public string TempleteWall { get; set; }
        public string TempleteFloor { get; set; }

        public Resources(string nameFileJornal, int heihgtMaxBox, string templeteWall, string templeteFloor)
        {
            NameFileJornal = nameFileJornal;
            HeihgtMaxBox = heihgtMaxBox;
            TempleteWall = templeteWall;
            TempleteFloor = templeteFloor;
        }
    }
}

using System;

namespace ExcelMacroAdd.Serializable
{
    [Serializable]
    public class AppSettings
    {
        public ResourcesForm1 ResourcesForm1 { get; set; }
        public ResourcesForm2 ResourcesForm2 { get; set; }
        public ResourcesDBConect ResourcesDBConect { get; set; }

        public AppSettings(ResourcesForm1 resourcesForm1, ResourcesForm2 resourcesForm2, ResourcesDBConect resourcesDBConect)
        {
            ResourcesForm1 = resourcesForm1;
            ResourcesForm2 = resourcesForm2;
            ResourcesDBConect = resourcesDBConect;          
        }
    }
}

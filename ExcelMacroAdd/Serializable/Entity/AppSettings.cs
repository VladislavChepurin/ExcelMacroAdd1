using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public Resources Resources { get; set; }
        public ResourcesForm2 ResourcesForm2 { get; set; }
        public ResourcesForm4 ResourcesForm4 { get; set; }

        public AppSettings(Resources resources, ResourcesForm2 resourcesForm2, ResourcesForm4 resourcesForm4)
        {
            Resources = resources;
            ResourcesForm2 = resourcesForm2;
            ResourcesForm4 = resourcesForm4;
        }
    }
}

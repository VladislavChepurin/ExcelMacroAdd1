using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public Resources Resources { get; set; }
        public ResourcesForm2 ResourcesForm2 { get; set; }
        public AppSettings(Resources resources, ResourcesForm2 resourcesForm2)
        {
            Resources = resources;
            ResourcesForm2 = resourcesForm2;      
        }
    }
}

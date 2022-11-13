using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public Resources Resources { get; set; }

        public  CorrectFontResources CorrectFontResources { get; set; }
        public string GlobalDateBaseLocation { get; set; }

        public AppSettings(Resources resources, CorrectFontResources correctFontResources, string globalDateBaseLocation)
        {
            Resources = resources;
            CorrectFontResources = correctFontResources;
            GlobalDateBaseLocation = globalDateBaseLocation;
        }
    }
}

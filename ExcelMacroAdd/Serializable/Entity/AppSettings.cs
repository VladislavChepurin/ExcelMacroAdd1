using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public FillingOutThePassportSettings Resources { get; set; }
        public  CorrectFontSettings CorrectFontResources { get; set; }
        public FormSettings FormSettings { get; set; }
        public string GlobalDateBaseLocation { get; set; }       

        public AppSettings(FillingOutThePassportSettings resources, CorrectFontSettings correctFontResources, FormSettings formSettings, string globalDateBaseLocation)
        {
            Resources = resources;
            CorrectFontResources = correctFontResources;
            FormSettings = formSettings;
            GlobalDateBaseLocation = globalDateBaseLocation;
        }
    }
}

using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public string LineseKey { get; set; }
        public FillingOutThePassportSettings Resources { get; set; }
        public CorrectFontSettings CorrectFontResources { get; set; }
        public FormSettings FormSettings { get; set; }
        public string GlobalDateBaseLocation { get; set; }
        public bool GlobalDateBaseLocationEnable { get; set; }
        public TypeNkySettings[] TypeNkySettings { get; set; }

        public AppSettings(string lineseKey, FillingOutThePassportSettings resources, CorrectFontSettings correctFontResources, FormSettings formSettings, string globalDateBaseLocation, TypeNkySettings[] typeNkySettings)
        {
            LineseKey = lineseKey;
            Resources = resources;
            CorrectFontResources = correctFontResources;
            FormSettings = formSettings;
            GlobalDateBaseLocation = globalDateBaseLocation;
            TypeNkySettings = typeNkySettings;
            LineseKey = lineseKey;
        }
    }
}

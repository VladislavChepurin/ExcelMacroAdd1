using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {      
        public FillingOutThePassportSettings Resources { get; set; }
        public CorrectFontSettings CorrectFontResources { get; set; }
        public FormSettings FormSettings { get; set; }
        public string GlobalDateBaseLocation { get; set; }
        public bool GlobalDateBaseLocationEnable { get; set; }
        public TypeNkySettings[] TypeNkySettings { get; set; }      
    }
}

using System;

namespace ExcelMacroAdd.Serializable
{
    [Serializable]
    public class AppSettings
    {
        public StringResourcesForm2 StringResourcesForm2 { get; set; }
        public StringResourcesMainRibbon StringResourcesMainRibbon { get; set; }

        public AppSettings(StringResourcesForm2 stringResourcesForm2, StringResourcesMainRibbon stringResourcesMainRibbon)
        {
            StringResourcesForm2 = stringResourcesForm2;
            StringResourcesMainRibbon = stringResourcesMainRibbon;
        }
    }
}

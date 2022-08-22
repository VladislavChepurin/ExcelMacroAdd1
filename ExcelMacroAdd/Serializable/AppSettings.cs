using System;
using System.Text.Json.Serialization;

namespace ExcelMacroAdd.Serializable
{
    [Serializable]
    public class AppSettings
    {
        StringResourcesForm2 StringResourcesForm2 { get; set; }
        StringResourcesMainRibbon StringResourcesMainRibbon { get; set; }

        public AppSettings(StringResourcesForm2 stringResourcesForm2, StringResourcesMainRibbon stringResourcesMainRibbon)
        {
            StringResourcesForm2 = stringResourcesForm2;
            StringResourcesMainRibbon = stringResourcesMainRibbon;
        }
    }
}

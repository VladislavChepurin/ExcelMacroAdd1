using System;

namespace ExcelMacroAdd.Serializable
{
    [Serializable]
    public class StringResourcesMainRibbon
    {
        public string ProviderData { get; set; }
        public string NameFileDB { get; set; }
        public string RealseDirectoryDB { get; set; }
        public string DebugDirectoryDB { get; set; }

        public StringResourcesMainRibbon(string providerData, string nameFileDB, string realseDirectoryDB, string debugDirectoryDB)
        {
            ProviderData = providerData;
            NameFileDB = nameFileDB;
            RealseDirectoryDB = realseDirectoryDB;
            DebugDirectoryDB = debugDirectoryDB;
        }
    }
}

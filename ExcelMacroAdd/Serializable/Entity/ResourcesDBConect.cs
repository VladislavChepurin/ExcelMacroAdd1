using ExcelMacroAdd.Interfaces;
using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class ResourcesDBConect : IResourcesDBConect
    {
        public string ProviderData { get; set; }
        public string NameFileDB { get; set; }

        public ResourcesDBConect(string providerData, string nameFileDB)
        {
            ProviderData = providerData;
            NameFileDB = nameFileDB;
        }
    }
}

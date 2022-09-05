using Newtonsoft.Json;
using System.IO;

namespace ExcelMacroAdd.Serializable
{
    public class AppSettingsDeserialize
    {
        private readonly string _jsonPatch;

        public AppSettingsDeserialize(string jsonPatch)
        {
            _jsonPatch = jsonPatch;
        }
            
        public AppSettings GetSettingsModels()
        {           
            JsonSerializer serializer = new JsonSerializer();
            using (StreamReader sw = new StreamReader(_jsonPatch))
            {
                using (JsonReader reader = new JsonTextReader(sw))
                {
                    return serializer.Deserialize<AppSettings>(reader);
                }
            }              
        }
    }
}

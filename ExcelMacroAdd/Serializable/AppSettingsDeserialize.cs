using ExcelMacroAdd.Serializable.Entity;
using ExcelMacroAdd.UserException;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;

namespace ExcelMacroAdd.Serializable
{
    public class AppSettingsDeserialize
    {
        private readonly string jsonPath;

        public AppSettingsDeserialize(string jsonPatch)
        {
            this.jsonPath = jsonPatch;
        }

        public AppSettings GetSettingsModels()
        {
            const int bufferSize = 4096; // Оптимальный размер буфера для файловых операций

            var settings = new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Error,
                NullValueHandling = NullValueHandling.Include,
                DateParseHandling = DateParseHandling.DateTimeOffset,
                FloatParseHandling = FloatParseHandling.Decimal
            };

            try
            {
                JsonSerializer serializer = new JsonSerializer();
                using (StreamReader sw = new StreamReader(this.jsonPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: bufferSize))
                {
                    using (JsonReader reader = new JsonTextReader(sw))
                    {
                        return JsonSerializer.CreateDefault(settings)
                            .Deserialize<AppSettings>(reader);
                    }
                }
            }
            catch (JsonException ex)
            {
                throw new SettingsLoadException("Ошибка формата JSON", ex);
            }
            catch (IOException ex)
            {
                throw new SettingsLoadException("Ошибка доступа к файлу", ex);
            }
            catch (Exception ex)
            {
                throw new SettingsLoadException("Неизвестная ошибка", ex);
            }
        }
    }
}

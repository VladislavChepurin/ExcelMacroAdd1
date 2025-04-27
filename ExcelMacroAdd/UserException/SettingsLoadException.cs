using ExcelMacroAdd.Services;
using System;

namespace ExcelMacroAdd.UserException
{
    public class SettingsLoadException: Exception
    {
        public SettingsLoadException(string message, Exception exception) : base(message)
        {
            Logger.LogException(exception, message);
        }
    }
}

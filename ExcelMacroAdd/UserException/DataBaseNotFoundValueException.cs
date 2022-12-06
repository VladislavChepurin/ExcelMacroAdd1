using System;

namespace ExcelMacroAdd.UserException
{
    public class DataBaseNotFoundValueException : Exception
    {
        public DataBaseNotFoundValueException(string message) : base(message)
        {
        }
    }
}

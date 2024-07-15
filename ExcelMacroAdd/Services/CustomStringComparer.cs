using System.Collections.Generic;

namespace ExcelMacroAdd.Services
{
    internal class CustomStringComparer : IEqualityComparer<string>
    {
        public bool Equals(string x, string y)
        {
            if (x is null || y is null) return false;
            return x.ToLower() == y.ToLower();

        }
        public int GetHashCode(string obj) => obj.ToLower().GetHashCode();
    }
}

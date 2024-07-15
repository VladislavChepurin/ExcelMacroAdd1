using ExcelMacroAdd.Models.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMacroAdd.Models
{
    internal class Dkc : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Dkc()
        {
            RangeSearch = new[] { "DKC", "ДКС" };
            OutValue = "Dkc";
        }
    }
}

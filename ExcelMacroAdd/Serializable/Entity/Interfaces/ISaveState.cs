using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMacroAdd.Serializable.Entity.Interfaces
{
    internal interface ISaveState
    {
        bool SaveWorkBook { get; set; }
        bool SaveWorkSheet { get; set; }
    }
}

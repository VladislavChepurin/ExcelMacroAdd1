using System.Collections.Generic;
using System.Linq;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IForm4Data
    {
        string[] GetComboBox2Items(string current);
        string[] GetComboBox3Items(string current, string bus);
        string[] GetComboBox4Items(string current, string bus, string accuracy);
    }
}

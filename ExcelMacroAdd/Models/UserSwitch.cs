using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    public class UserSwitch: IUserSwitch
    {
        public string group { get; set; }
        public int[] current {  get; set; }
        public string[] quantityPole { get; set; }

        public UserSwitch(string group, int[] current, string[] quantityPole)
        {
            this.group = group;
            this.current = current;
            this.quantityPole = quantityPole;
        }
    }
}

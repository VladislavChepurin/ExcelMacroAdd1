using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    public class UserCircuitBreaker: IUserCircuitBreaker
    {
        public string group { get; set; }
        public int[] current {  get; set; }
        public string[] kurve { get; set; }
        public string[] maxCurrent { get; set; }
        public string[] quantityPole { get; set; }

        public UserCircuitBreaker(string group, int[] current, string[] kurve, string[] maxCurrent, string[] quantityPole)
        {
            this.group = group;
            this.current = current;
            this.kurve = kurve;
            this.maxCurrent = maxCurrent;
            this.quantityPole = quantityPole;
        }
    }
}

using ExcelMacroAdd.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBaseTestConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (DataContext db = new DataContext())
            {
                var users = db.JornalNKU;
                foreach (JornalNKU u in users)
                {
                    Console.WriteLine($"{u.Id} ---> {u.Article}");
                }
            }
        }
    }
}

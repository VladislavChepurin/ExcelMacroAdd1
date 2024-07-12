using System;
using System.Diagnostics;

namespace ExcelMacroAdd.Functions
{
    internal class UpdatingCalculation : AbstractFunctions
    {
        public override void Start()
        {
            int currentRow = 2;
            int counter = 0;

            while (true)
            {
                if (Worksheet.Range["A" + currentRow].Value == null && Worksheet.Range["B" + currentRow].Value == null)
                {
                    if (counter > 1)
                    {
                        break;
                    }
                    counter++;
                }
                else
                {
                    counter = 0;
                }
               




                currentRow++;

           
            }



            
        }
    }
}

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculate_Formula_once
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "Without creating formulae chain.xls";

            //Load the template workbook
            Workbook workbook = new Workbook(filePath);

            //Print the time before formula calculation
            Console.WriteLine(DateTime.Now);

            //Set the CreateCalcChain as false
            workbook.Settings.CreateCalcChain = false;

            //Calculate the workbook formulas
            workbook.CalculateFormula();

            //Print the time after formula calculation
            Console.WriteLine(DateTime.Now);
            workbook.Save("Without creating formulae chain.xls");
        }
    }
}

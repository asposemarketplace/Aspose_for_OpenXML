using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_from_Workshet
{
    class Program
    {
        static void Main(string[] args)
        {
        }
        public static void ColumnsContainingStronglyTypedData()
        {
            string MyDir = "";
            //Creating a file stream containing the Excel file to be opened
            FileStream fstream = new FileStream(MyDir+"Export File.xls", FileMode.Open);

            //Instantiating a Workbook object
            //Opening the Excel file through the file stream
            Workbook workbook = new Workbook(fstream);

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[0];

            //Exporting the contents of 1 rows and 2 columns starting from 1st cell to DataTable
            DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, 1, 2, true);

            workbook.Save(MyDir + "Exported Data.xls");
            //Closing the file stream to free all resources
            fstream.Close();
 
        }
        public static void ColumnsContainingNonStronglyTypedData()
        {

        }
    }
}
    
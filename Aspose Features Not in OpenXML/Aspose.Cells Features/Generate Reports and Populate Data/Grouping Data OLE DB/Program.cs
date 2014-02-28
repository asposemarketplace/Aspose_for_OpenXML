using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grouping_Data_OLE_DB
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Create a connection object, specify the provider info and set the data source.
            OleDbConnection con = new OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=F:\\Dropbox\\Zeeshan\\Aspose Vs OpenXML\\Features Supported by Aspose not Open XML\\Aspose.Cells Features\\Generate Reports and Populate Data\\Grouping Data OLE DB\\Data\\Northwind.mdb");

            //Open the connection object.
            con.Open();

            //Create a command object and specify the SQL query.
            OleDbCommand cmd = new OleDbCommand("Select * from [Order Details]", con);

            //Create a data adapter object.
            OleDbDataAdapter da = new OleDbDataAdapter();

            //Specify the command.
            da.SelectCommand = cmd;

            //Create a dataset object.
            DataSet ds = new DataSet();

            //Fill the dataset with the table records.
            da.Fill(ds, "Order Details");

            //Create a datatable with respect to dataset table.
            DataTable dt = ds.Tables["Order Details"];

            //Create WorkbookDesigner object.
            WorkbookDesigner wd = new WorkbookDesigner();

            //Open the template file (which contains smart markers).
            wd.Workbook = new Workbook(MyDir+"SmartMarkerDesigner.xls");

            //Set the datatable as the data source.
            wd.SetDataSource(dt);

            //Process the smart markers to fill the data into the worksheets.
            wd.Process(true);

            //Save the excel file.
            wd.Workbook.Save(MyDir+"outSmartMarker_Designer.xls");
 
        }
    }
}

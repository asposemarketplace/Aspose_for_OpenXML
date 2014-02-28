using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Field_Update
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            Document doc = new Document(MyDir + "Rendering.docx");

            // This updates all fields in the document.
            doc.UpdateFields();

            doc.Save(MyDir + "Rendering.UpdateFields Out.pdf");
        }
    }
}

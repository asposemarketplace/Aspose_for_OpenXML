using Aspose.Slides;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converting_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir ="";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir+"Conversion.ppt");
                //Save the presentation to PDF with default options
            pres.Save(MyDir + "Converted.pdf", Aspose.Slides.Export.SaveFormat.Pdf);
        }
    }
}

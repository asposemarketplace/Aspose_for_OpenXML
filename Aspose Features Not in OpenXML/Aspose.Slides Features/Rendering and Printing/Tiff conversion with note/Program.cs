using Aspose.Slides.Export;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tiff_conversion_with_note
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.pptx");

            //Saving the presentation to TIFF notes
            pres.Save(MyDir + "Converted with Notes.tiff", SaveFormat.TiffNotes);
        }

    }
}

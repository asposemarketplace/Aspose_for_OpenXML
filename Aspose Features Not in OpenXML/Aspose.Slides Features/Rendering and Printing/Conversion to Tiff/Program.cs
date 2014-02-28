using Aspose.Slides;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Conversion_to_Tiff
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Instantiate a Presentation object that represents a presentation file
            using (PresentationEx pres = new PresentationEx(MyDir+"Conversion.pptx"))
            {

                //Saving the presentation to TIFF document
                pres.Save(MyDir + "Converted.tiff", Aspose.Slides.Export.SaveFormat.Tiff);
            }
        }
    }
}

using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Pptx;
namespace Converting_to_XPS
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");
                //Saving the presentation to TIFF document
            pres.Save(MyDir + "converted.xps", Aspose.Slides.Export.SaveFormat.Xps);
        }
    }
}

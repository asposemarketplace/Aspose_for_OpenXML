using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Conversion_PPT_to_PPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir="";
            //Instantiate a Presentation object that represents a PPTX file
            PresentationEx pres = new PresentationEx(MyDir+"Conversion.ppt");
            //Saving the PPTX presentation to PPTX format
            pres.Save(MyDir +"Converted.pptx", SaveFormat.Pptx);
        }
    }
}

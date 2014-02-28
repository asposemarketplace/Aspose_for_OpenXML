using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converting_to_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");

                HtmlOptions htmlOpt = new HtmlOptions();
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

                //Saving the presentation to HTML
                pres.Save(MyDir + "Converted.html", Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
         }
    }
}

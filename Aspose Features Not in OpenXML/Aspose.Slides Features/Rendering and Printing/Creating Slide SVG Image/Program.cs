using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creating_Slide_SVG_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate a Presentation class that represents the presentation file
            string MyDir = "";
            using (PresentationEx pres = new PresentationEx(MyDir+"Slides Test Presentation.pptx"))
            {

                //Access the second slide
                SlideEx sld = pres.Slides[1];

                //Create a memory stream object
                MemoryStream SvgStream = new MemoryStream();

                //Generate SVG image of slide and save in memory stream
                sld.WriteAsSvg(SvgStream);
                SvgStream.Position = 0;

                //Save memory stream to file
                using (Stream fileStream = System.IO.File.OpenWrite(MyDir+ "PresentatoinTemplate.svg"))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = SvgStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        fileStream.Write(buffer, 0, len);
                    }

                }
                SvgStream.Close();
            }
 
        }
    }
}

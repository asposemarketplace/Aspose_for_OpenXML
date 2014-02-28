using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Slide_Thumbnail_to_JPEG
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            //Instantiate a Presentation class that represents the presentation file
            using (PresentationEx pres = new PresentationEx(MyDir+"Slides Test Presentation.pptx"))
            {

                //Access the first slide
                SlideEx sld = pres.Slides[0];

                //Create a full scale image
                Bitmap bmp = sld.GetThumbnail(1f, 1f);

                //Save the image to disk in JPEG format
                bmp.Save(MyDir+"Test Thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            }
        }
    }
}

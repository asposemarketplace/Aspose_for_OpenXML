using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Render_doc_in_Pages_as_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Open the document.
            Document doc = new Document(dataDir + "SaveAsPNG.doc");
            
            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.Resolution = 160;

            // Save each page of the document as Png.
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(string.Format(dataDir+i+"SaveAsPNG out.Png", i), options);
            }

        }
    }
}

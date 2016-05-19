﻿// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Conversion.pptx";
            string destFileName = FilePath + "Converting to HTML.html";
            
            //Instantiate a Presentation object that represents a presentation file
            Presentation pres = new Presentation(srcFileName);

            HtmlOptions htmlOpt = new HtmlOptions();
            htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

            //Saving the presentation to HTML
            pres.Save(destFileName, Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
        }
    }
}

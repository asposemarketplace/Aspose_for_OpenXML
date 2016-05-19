﻿// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintByDefaultPrinter();
            PrintBySpecificPrinter();
        }
        public static void PrintByDefaultPrinter()
        {
            string MyDir = @"..\..\..\Sample Files\";
            //Load the presentation
            Presentation asposePresentation = new Presentation(MyDir + "Print.pptx");

            //Call the print method to print whole presentation to the default printer
            asposePresentation.Print();

        }
        public static void PrintBySpecificPrinter()
        {
            string MyDir = @"..\..\..\Sample Files\";
            //Load the presentation
            Presentation asposePresentation = new Presentation(MyDir + "Print.pptx");

            //Call the print method to print whole presentation to the desired printer
            asposePresentation.Print("LaserJet1100");

        }
    }
}

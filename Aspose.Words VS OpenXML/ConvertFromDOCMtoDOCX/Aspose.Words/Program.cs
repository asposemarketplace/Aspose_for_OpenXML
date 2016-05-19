﻿using Aspose.Words;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "ConvertFromDOCMtoDOCX - Aspose.docm";
            string NewFile = FilePath + "ConvertFromDOCMtoDOCX - Aspose - Output.docx";
            ConvertDOCMtoDOCX(File, NewFile);
        }

        private static void ConvertDOCMtoDOCX(string oldfileName, string newfilename)
        {
            Document doc = new Document(oldfileName);
            doc.Save(newfilename, SaveFormat.Docx);
        }
    }
}

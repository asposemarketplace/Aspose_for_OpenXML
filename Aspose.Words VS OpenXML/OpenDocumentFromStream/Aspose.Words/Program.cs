﻿using Aspose.Words;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        private static string FilePath = @"..\..\..\..\Sample Files\";
        private static string FileName = FilePath + "OpenDocumentFromStream.docx";
            
        
        static void Main(string[] args)
        {
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            Stream stream = File.Open(FileName, FileMode.Open);
            OpenAndAddToWordprocessingStream(stream, txt);
        }

        private static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            Document doc = new Document(stream);
            stream.Close();
            DocumentBuilder db = new DocumentBuilder(doc);
            db.Writeln(txt);
            doc.Save(FileName);
        }
    }
}

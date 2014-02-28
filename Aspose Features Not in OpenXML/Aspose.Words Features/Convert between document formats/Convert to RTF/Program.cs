using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
namespace Convert_to_RTF
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            Document doc = new Document(MyDir + "Converting Document.docx");
            doc.Save(MyDir + "Document.Doc2RtfSave Out.rtf");
        }
    }
}

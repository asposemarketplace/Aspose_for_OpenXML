using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
namespace keping_the_content_frm_split
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = "";
            Document dstDoc = new Document(MyDir + "src.doc");
            Document srcDoc = new Document(MyDir+ "srcDocument.doc");

            // Set the source document to appear straight after the destination document's content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Iterate through all sections in the source document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(MyDir + "JoiningDocumentByKeepingContentfromSplit.doc");
        }
    }
}

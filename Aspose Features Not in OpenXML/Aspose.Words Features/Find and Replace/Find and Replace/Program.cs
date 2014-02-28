using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Find_and_Replace
{

    class Program
    {
        
        static void Main(string[] args)
        {
            string MyDir = "";
            Document doc = new Document(MyDir+"Find and Replace.doc");
            ReplaceOneWordWithAnother(doc);
            doc.Save(MyDir + "Range.ReplaceOneWordWithAnother Out.doc");
            ReplaceTwoSimilarWords(doc);
            doc.Save(MyDir + "Range.ReplaceTwoSimilarWords Out.doc");

        }
        static void ReplaceOneWordWithAnother(Document doc)
        {
            doc.Range.Replace("sad", "bad", false, true);
            
        }
        static void ReplaceTwoSimilarWords(Document doc)
        {
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");
            
        }
        public void ReplaceWithEvaluator()
        {
            string MyDir = "";
            Document doc = new Document(MyDir + "Range.ReplaceWithEvaluator.doc");
            doc.Range.Replace(new Regex("[s|m]ad"), new MyReplaceEvaluator(), true);
            doc.Save(MyDir + "Range.ReplaceWithEvaluator Out.doc");
        }

        private class MyReplaceEvaluator : IReplacingCallback
        {
            /// <summary>
            /// This is called during a replace operation each time a match is found.
            /// This method appends a number to the match string and returns it as a replacement string.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                e.Replacement = e.Match.ToString() + mMatchNumber.ToString();
                mMatchNumber++;
                return ReplaceAction.Replace;
            }

            private int mMatchNumber;
        }
 
    }
}

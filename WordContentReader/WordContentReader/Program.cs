using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;


class Program
{
    static void Main(string[] args)
    {
        // Specify the path of the Word document
        string wordFile = @"C:/Users/Romano/Desktop/TestWord.docx";

        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFile, false))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Get the text of the document
            string text = mainPart.Document.InnerText;

            // Specify the path of the text file
            string textFile = @"C:/Users/Romano/Desktop/TestOutput.txt";

            // Write the text to the file
            File.WriteAllText(textFile, text);
        }
    }
}
using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        //ReadOnlyTableContent();
        //ReadEachTableContent();
        //ReadAndWriteWordDataToTextfile()
    }

    /// <summary>
    /// Reads only content from the tables in word document
    /// </summary>
    static void ReadOnlyTableContent()
    {
        // Specify the path of the Word document
        string wordFile = @"C:\Users\Romano\Desktop\TestWord.docx";

        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFile, false))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;
            var tables = mainPart.Document.Body.Elements<Table>();

            // Specify the path of the text file
            string textFile = @"C:\Users\Romano\Desktop\TestOutput.txt";

            // Iterate through each table
            foreach (Table table in tables)
            {
                var rows = table.Elements<TableRow>();

                // Iterate through each row
                foreach (TableRow row in rows)
                {
                    var cells = row.Elements<TableCell>();

                    // Iterate through each cell
                    foreach (TableCell cell in cells)
                    {
                        var paragraphs = cell.Elements<Paragraph>();
                        // Iterate through each paragraph
                        foreach (Paragraph paragraph in paragraphs)
                        {
                            var runs = paragraph.Elements<Run>();

                            // Iterate through each run
                            foreach (Run run in runs)
                            {
                                // Get the text of the run
                                var text = run.InnerText;
                                // Write the text to the file, followed by a new line
                                File.AppendAllText(textFile, text + "\r\n");
                            }
                        }
                    }
                }
                // Write a new line after each table
                File.AppendAllText(textFile, "\r\n");
            }
        }
    }

    /// <summary>
    /// Read each table in the document
    /// Iterate through each table row and cell and write it in the text file
    /// </summary>
    static void ReadEachTableContent()
    {
        // Specify the path of the Word document
        string wordFile = @"C:\Users\Romano\Desktop\TestWord.docx";

        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFile, false))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;
            var tables = mainPart.Document.Body.Elements<Table>();

            // Specify the path of the text file
            string textFile = @"C:\Users\Romano\Desktop\TestOutput.txt";

            // Iterate through each table
            foreach (Table table in tables)
            {
                var rows = table.Elements<TableRow>();

                // Iterate through each row
                foreach (TableRow row in rows)
                {
                    var cells = row.Elements<TableCell>();

                    // Iterate through each cell
                    foreach (TableCell cell in cells)
                    {
                        var text = cell.InnerText;
                        // Write the text to the file
                        File.AppendAllText(textFile, text + "\r\n");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Just write everything from the document into the text file
    /// </summary>
    static void ReadAndWriteWordDataToTextfile()
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

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

class Program
{
    static void Main(string[] args)
    {
        //ReadOnlyTableContentOnSteroids();
        ReadTableAndWriteContentOnSteroids();
        //ReadOnlyTableContent();
        //ReadEachTableContent();
        //ReadAndWriteWordDataToTextfile()
    }

    /// <summary>
    /// Check if the current cell is the last one in the row or not
    /// If it's not the last one it will append a comma to separate the content of the cell to the next one
    /// If it's the last one it will append only the text
    /// </summary>
    static void ReadTableAndWriteContentOnSteroids()
    {
        // Specify the path of the Word document
        string wordFile = @"C:/Users/Romano/Desktop/TestWord.docx";
        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFile, false))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;
            var tables = mainPart.Document.Body.Elements<Table>();

            // Specify the path of the text file
            string textFile = @"C:\Users\Romano\Desktop\TestOutput.txt";

            KillWordAndTextfileProcess("", textFile);
            // Rewrite the text file each time the program is run
            File.WriteAllText(textFile, string.Empty);

            // Iterate through each table
            for (int i = 0; i < tables.Count(); i++)
            {
                var rows = tables.ElementAt(i).Elements<TableRow>();

                // Iterate through each row
                for (int j = 0; j < rows.Count(); j++)
                {
                    var cells = rows.ElementAt(j).Elements<TableCell>();

                    // Iterate through each cell
                    for (int k = 0; k < cells.Count(); k++)
                    {
                        var paragraphs = cells.ElementAt(k).Elements<Paragraph>();
                        // Iterate through each paragraph
                        for (int l = 0; l < paragraphs.Count(); l++)
                        {
                            var runs = paragraphs.ElementAt(l).Elements<Run>();

                            // Iterate through each run
                            for (int m = 0; m < runs.Count(); m++)
                            {
                                // Get the text of the run
                                var text = runs.ElementAt(m).InnerText;
                                // Write the text to the file, followed by a comma if it's not the last cell in the row
                                if (k != cells.Count() - 1)
                                {
                                    File.AppendAllText(textFile, text + ", ");
                                }
                                else
                                {
                                    File.AppendAllText(textFile, text);
                                }
                            }
                        }
                    }
                    // Write a new line after each row
                    File.AppendAllText(textFile, "\r\n");
                }
                // Write a new line after each table
                File.AppendAllText(textFile, "\r\n");
            }
            System.Diagnostics.Process.Start("explorer.exe", textFile);
        }
    }

    static void KillWordAndTextfileProcess(string wordFile, string textFile)
    {
        // Kill any processes associated with the Word document
        System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
        foreach (System.Diagnostics.Process process in processes)
        {
            process.Kill();
        }
        // Delete the Word document
        //File.Delete(wordFile);
        // Kill any processes associated with the text file
        processes = System.Diagnostics.Process.GetProcessesByName("notepad");
        foreach (System.Diagnostics.Process process in processes)
        {
            process.Kill();
        }
        // Delete the text file
        File.Delete(textFile);
    }

    static void ReadOnlyTableContentOnSteroids()
    {
        // Specify the path of the Word document
        string wordFile = @"C:/Users/Romano/Desktop/TestWord.docx";

        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFile, false))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;
            var tables = mainPart.Document.Body.Elements<Table>();

            // Specify the path of the text file
            string textFile = @"C:\Users\Romano\Desktop\TestOutput.txt";

            // Iterate through each table
            for (int i = 0; i < tables.Count(); i++)
            {
                var rows = tables.ElementAt(i).Elements<TableRow>();

                // Iterate through each row
                for (int j = 0; j < rows.Count(); j++)
                {
                    var cells = rows.ElementAt(j).Elements<TableCell>();

                    // Iterate through each cell
                    for (int k = 0; k < cells.Count(); k++)
                    {
                        var paragraphs = cells.ElementAt(k).Elements<Paragraph>();
                        // Iterate through each paragraph
                        for (int l = 0; l < paragraphs.Count(); l++)
                        {
                            var runs = paragraphs.ElementAt(l).Elements<Run>();

                            // Iterate through each run
                            for (int m = 0; m < runs.Count(); m++)
                            {
                                // Get the text of the run
                                var text = runs.ElementAt(m).InnerText;
                                File.AppendAllText(textFile, text + "\r\n");
                            }
                        }
                    }
                }
            }
        }
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

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.IO;
using System.Linq;

public class Program
{
    public string ExtractTextFromDocx(string filename)
    {
        var sb = new StringBuilder();

        using (var doc = WordprocessingDocument.Open(filename, false))
        {
            var body = doc.MainDocumentPart?.Document.Body;
            if (body == null) return string.Empty;

            foreach (var element in body.Elements())
            {
                // Handle Paragraphs
                if (element is Paragraph paragraph)
                {
                    foreach (var text in paragraph.Descendants<Text>())
                    {
                        sb.AppendLine(text.Text);
                    }
                    sb.AppendLine();
                }
                // Handle Tables
                else if (element is Table table)
                {
                    sb.AppendLine(ConvertTableToMarkdown(table));
                    sb.AppendLine();
                }
            }
        }

        return sb.ToString();
    }

    private string ConvertTableToMarkdown(Table table)
    {
        var sb = new StringBuilder();

        // Convert table rows to markdown format
        foreach (var row in table.Elements<TableRow>())
        {
            var rowTexts = row.Elements<TableCell>()
                              .Select(cell => GetCellText(cell))
                              .ToArray();

            // Markdown table row
            sb.AppendLine("| " + string.Join(" | ", rowTexts) + " |");

            // Add header delimiter after the first row
            if (row == table.Elements<TableRow>().First())
            {
                sb.AppendLine("|" + string.Join("|", rowTexts.Select(_ => "---")) + "|");
            }
        }

        return sb.ToString();
    }

    private string GetCellText(TableCell cell)
    {
        // Extract all text within the table cell
        var cellText = new StringBuilder();
        foreach (var text in cell.Descendants<Text>())
        {
            cellText.Append(text.Text);
        }
        return cellText.ToString().Trim();  // Trim to remove any unnecessary whitespaces
    }

    public static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: und <filename>");
            return;
        }

        if (!File.Exists(args[0]))
        {
            Console.WriteLine("File not found: " + args[0]);
            return;
        }

        var program = new Program();
        var text = program.ExtractTextFromDocx(args[0]);
        Console.WriteLine(text);
    }
}

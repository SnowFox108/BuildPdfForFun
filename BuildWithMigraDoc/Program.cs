// See https://aka.ms/new-console-template for more information


using BuildWithMigraDoc;using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

Console.WriteLine("Hello, World!");

new DocumentBuilder();

Console.WriteLine("Press any key to continue...");

Console.ReadKey();
return;

// Create a MigraDoc document
Document document = CreateDocument();
document.UseCmykColor = true;

// ===== Unicode encoding and font program embedding in MigraDoc is demonstrated here =====

// A flag indicating whether to create a Unicode PDF or a WinAnsi PDF file.
// This setting applies to all fonts used in the PDF document.
// This setting has no effect on the RTF renderer.
const bool unicode = true;

// An enum indicating whether to embed fonts or not.
// This setting applies to all font programs used in the document.
// This setting has no effect on the RTF renderer.
// (The term 'font program' is used by Adobe for a file containing a font. Technically a 'font file'
// is a collection of small programs and each program renders the glyph of a character when executed.
// Using a font in PDFsharp may lead to the embedding of one or more font programms, because each outline
// (regular, bold, italic, bold+italic, ...) has its own fontprogram)
const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

// ========================================================================================

// Create a renderer for the MigraDoc document.
var pdfRenderer = new PdfDocumentRenderer(unicode);

// Associate the MigraDoc document with a renderer
pdfRenderer.Document = document;
const string filename = @"..\..\..\..\SupportFiles\HelloWorld.pdf";

// Layout and render document to PDF
pdfRenderer.RenderDocument();

// Save the document...
pdfRenderer.PdfDocument.Save(filename);
// ...and start a viewer.
//Process.Start(filename);



static Document CreateDocument()
{
    // Create a new MigraDoc document
    Document document = new Document();
    document.Info.Title = "Hello World";

    // Add a section to the document
    Section section = document.AddSection();

    // Add a paragraph to the section
    Paragraph paragraph = section.AddParagraph();

    paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);

    var text =
        @"Loboreet autpat, C./ Núñez de Balboa 35-A, 3rd floor B quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet, consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tat laor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";

    // Add some text to the paragraph
    paragraph.AddText(text);
    //paragraph.AddFormattedText("Hello, World!", TextFormat.Bold);

    for (int i = 0; i < 15; i++)
        AddParagraph(document, text);

    DemonstrateSimpleTable(document);
    return document;
}

static void AddParagraph(Document document, string text)
{
    document.LastSection.AddParagraph("Justified", "Heading3");

    Paragraph paragraph = document.LastSection.AddParagraph();
    paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
    paragraph.Format.Alignment = ParagraphAlignment.Justify;
    paragraph.AddText(text);
}

static void DemonstrateSimpleTable(Document document)
{
    document.LastSection.AddParagraph("Simple Tables", "Heading2");

    Table table = new Table();
    table.Borders.Width = 0.75;

    Column column = table.AddColumn(Unit.FromCentimeter(2));
    column.Format.Alignment = ParagraphAlignment.Center;

    table.AddColumn(Unit.FromCentimeter(5));

    Row row = table.AddRow();
    row.Shading.Color = Colors.PaleGoldenrod;
    Cell cell = row.Cells[0];
    cell.AddParagraph("Itemus");
    cell = row.Cells[1];
    cell.AddParagraph("Descriptum");

    row = table.AddRow();
    cell = row.Cells[0];
    cell.AddParagraph("1");
    cell = row.Cells[1];
    cell.AddParagraph(FillerText.ShortText);

    row = table.AddRow();
    cell = row.Cells[0];
    cell.AddParagraph("2");
    cell = row.Cells[1];
    cell.AddParagraph(FillerText.Text);

    table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

    document.LastSection.Add(table);
}

public class FillerText
{
    public static string Text = @"Loboreet autpat, C./ Núñez de Balboa 35-A, 3rd floor B quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet, consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tat laor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
    public static string ShortText = @"Núñez de Balboa 35-A";
}
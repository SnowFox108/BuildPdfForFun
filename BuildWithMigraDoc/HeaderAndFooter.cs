
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Shapes;
using MigraDocCore.Rendering;

namespace BuildWithMigraDoc;

public class HeaderAndFooter
{
    private Document _document;
    private Section _section;
    private TextFrame _leftTextFrame;
    private TextFrame _rightTextFrame;

    public HeaderAndFooter()
    {
        _document = new Document();
        _document.Info.Title = "A sample document";
        _document.Info.Subject = "Demonstration";
        _document.Info.Author = "Tester";

        DefineStyles();
        CreatePage();
        FillContent();

        var pdfRenderer = new PdfDocumentRenderer(true);

        pdfRenderer.Document = _document;

        // Layout and render document to PDF
        pdfRenderer.RenderDocument();

        // Save the document...
        pdfRenderer.PdfDocument.Save(FillerText.Filename);
    }

    private void DefineStyles()
    {
        // Get the predefined style Normal.
        var style = _document.Styles["Normal"];
        // Because all styles are derived from Normal, the next line changes the 
        // font of the whole document. Or, more exactly, it changes the font of
        // all styles and paragraphs that do not redefine the font.
        style.Font.Name = "Segoe UI";

        style = _document.Styles[StyleNames.Header];
        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

        style = _document.Styles[StyleNames.Footer];
        style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

        // Create a new style called Table based on style Normal.
        style = _document.Styles.AddStyle("Table", "Normal");
        style.Font.Name = "Segoe UI Semilight";
        style.Font.Size = 9;

        // Create a new style called Title based on style Normal.
        style = _document.Styles.AddStyle("Title", "Normal");
        style.Font.Name = "Segoe UI Semibold";
        style.Font.Size = 9;

        // Create a new style called Reference based on style Normal.
        style = _document.Styles.AddStyle("Reference", "Normal");
        style.ParagraphFormat.SpaceBefore = "5mm";
        style.ParagraphFormat.SpaceAfter = "5mm";
        style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
    }

    private void CreatePage()
    {
        // Each MigraDoc document needs at least one section.
        _section = _document.AddSection();

        // Put a logo in the header

        Paragraph paragraph = _section.Headers.Primary.AddParagraph();
        paragraph.AddText("This is header");
        paragraph.Format.Font.Size = 9;

        // Create footer

        paragraph = _section.Footers.Primary.AddParagraph();
        paragraph.AddText("PowerBooks Inc · Sample Street 42 · 56789 Cologne · Germany");
        paragraph.Format.Font.Size = 9;
        paragraph.Format.Alignment = ParagraphAlignment.Center;

        // Create the text frame for the address
        _leftTextFrame = _section.AddTextFrame();
        _leftTextFrame.Height = "3.0cm";
        _leftTextFrame.Width = "7.0cm";
        _leftTextFrame.Left = ShapePosition.Left;
        _leftTextFrame.RelativeHorizontal = RelativeHorizontal.Margin;
        _leftTextFrame.Top = "5.0cm";
        _leftTextFrame.RelativeVertical = RelativeVertical.Page;

        _rightTextFrame = _section.AddTextFrame();
        //_rightTextFrame.Height = "3.0cm";
        _rightTextFrame.Width = "10.0cm";
        _rightTextFrame.Left = "7.0cm";
        //_rightTextFrame.Left = ShapePosition.Left;
        _rightTextFrame.RelativeHorizontal = RelativeHorizontal.Margin;
        //_rightTextFrame.Top = "5.0cm";
        _rightTextFrame.RelativeVertical = RelativeVertical.Margin;

    }

    private void FillContent()
    {
        // Fill the address in the address text frame.
        var paragraph = _leftTextFrame.AddParagraph();
        paragraph.AddText("name/singleName");
        paragraph.AddLineBreak();
        paragraph.AddText("address/line1");
        paragraph.AddLineBreak();
        paragraph.AddText("address/postalCode" + " " + "address/city");

        for (int i = 0; i < 15; i++)
        {
            paragraph = _rightTextFrame.AddParagraph();

            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.Format.Alignment = ParagraphAlignment.Justify;
            paragraph.AddText(FillerText.Text);

        }

        for (int i = 0; i < 55; i++)
        {
            paragraph = _section.AddParagraph();

            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.Format.Alignment = ParagraphAlignment.Justify;
            paragraph.AddText(FillerText.Text);

        }


    }
}

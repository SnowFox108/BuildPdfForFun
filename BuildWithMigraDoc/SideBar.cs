using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDocCore.DocumentObjectModel;

namespace BuildWithMigraDoc;

public class SideBar
{
    private Document _document;
    private Section _section;

    public Document Document() => _document;

    public SideBar()
    {
        _document = new Document();
        _document.Info.Title = "A sample document Header";
        _document.Info.Subject = "Demonstration";
        _document.Info.Author = "Tester";

        DefineStyles();
        CreatePage();

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
        _section.PageSetup.PageWidth = Unit.FromCentimeter(4);
        _section.PageSetup.PageHeight = Unit.FromCentimeter(29);
        _section.PageSetup.LeftMargin = Unit.FromCentimeter(1);
        _section.PageSetup.RightMargin = Unit.Zero;
        // Put a logo in the header
        Paragraph paragraph = _section.AddParagraph();
        paragraph.AddText("");
        paragraph.Format.Font.Size = 18;

        paragraph = _section.AddParagraph();
        paragraph.AddText("");
        paragraph.Format.Font.Size = 18;

        paragraph = _section.AddParagraph();
        paragraph.AddText("");
        paragraph.Format.Font.Size = 14;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        paragraph = _section.AddParagraph();
        paragraph.AddText("This is Sidebar");
        paragraph.Format.Font.Size = 9;

        for (int i = 0; i < 5; i++)
        {
            paragraph = _section.AddParagraph();
            paragraph.AddText(FillerText.Text);
            paragraph.Format.Font.Size = 9;
        }

        //// Create footer

        //paragraph = _section.Footers.Primary.AddParagraph();
        //paragraph.AddText("PowerBooks Inc · Sample Street 42 · 56789 Cologne · Germany");
        //paragraph.Format.Font.Size = 9;
        //paragraph.Format.Alignment = ParagraphAlignment.Center;

    }

}


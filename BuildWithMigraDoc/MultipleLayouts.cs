using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Shapes;
using MigraDocCore.Rendering;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace BuildWithMigraDoc;

public class MultipleLayouts
{
    private PdfDocument _document;
    private Section _section;
    private TextFrame _leftTextFrame;
    private TextFrame _rightTextFrame;
    private Header _header;
    private Footer _footer;
    private HeaderAndFooter _headerAndFooterDoc;
    private SideBar _sideDoc;
    private TopBar _topBar;

    private readonly double A4Width = XUnit.FromCentimeter(21).Point;
    private readonly double A4Height = XUnit.FromCentimeter(29.7).Point;

    public MultipleLayouts()
    {
        _header = new Header();
        _footer = new Footer();
        _headerAndFooterDoc = new HeaderAndFooter();
        _sideDoc = new SideBar();
        _topBar = new TopBar();


        _document = new PdfDocument();
        _document.Info.Title = "A sample document";
        _document.Info.Subject = "Demonstration";
        _document.Info.Author = "Tester";

        BuildPage();
        BuildPdfSharpPage();
        BuildPage();

        // Save the document...
        _document.Save(FillerText.Filename);

    }

    private void BuildPage()
    {
        var page = _document.AddPage();
        var gfx = XGraphics.FromPdfPage(page);
        // HACK²
        gfx.MUH = PdfFontEncoding.Unicode;

        var font = new XFont("Segoe UI", 13, XFontStyle.Bold);

        gfx.DrawString("The following paragraph was rendered using MigraDoc:", font, XBrushes.Black,
            new XRect(100, 100, page.Width - 200, 300), XStringFormats.Center);

        // You always need a MigraDoc document for rendering.
        var doc = new Document();
         // Create a renderer and prepare (=layout) the document.
        var docRenderer = new DocumentRenderer(doc);
        docRenderer.PrepareDocument();

        var sec = doc.AddSection();

        for (int i = 0; i < 1; i++)
        {
            // Add a single paragraph with some text and format information.
            var para = sec.AddParagraph();
            para.Format.Alignment = ParagraphAlignment.Justify;
            para.Format.Font.Name = "Times New Roman";
            para.Format.Font.Size = 12;
            para.Format.Font.Color = Colors.DarkGray;
            para.AddText("Duisism odigna acipsum delesenisl ");
            para.AddFormattedText("ullum in velenit", TextFormat.Bold);
            para.AddText(FillerText.Text);
            para.Format.Borders.Distance = "5pt";
            para.Format.Borders.Color = Colors.Gold;

            // Render the paragraph. You can render tables or shapes the same way.
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(5), XUnit.FromCentimeter(10), "12cm", para);

        }

    }

    private void BuildPdfSharpPage()
    {
        var headerDoc = _header.Document();
        var headerRenderer = new DocumentRenderer(headerDoc);
        headerRenderer.PrepareDocument();

        var footerDoc = _footer.Document();
        var footerRenderer = new DocumentRenderer(footerDoc);
        footerRenderer.PrepareDocument();

        var topDoc = _topBar.Document();
        var topRenderer = new DocumentRenderer(topDoc);
        topRenderer.PrepareDocument();
        var pageHeight = topRenderer.FormattedDocument.GetCurrentMigraDocPosition();

        var doc = _headerAndFooterDoc.Document();
        var docRenderer = new DocumentRenderer(doc);
        docRenderer.PrepareDocument();
        var pageCount = docRenderer.FormattedDocument.PageCount;

        var sideDoc = _sideDoc.Document();
        var sideDocRenderer = new DocumentRenderer(sideDoc);
        sideDocRenderer.PrepareDocument();
        var sidePageCount = sideDocRenderer.FormattedDocument.PageCount;


        for (var idx = 0; idx < pageCount; idx++)
        {
            var page = _document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            // HACK²
            gfx.MUH = PdfFontEncoding.Unicode;


            // Create document from HelloMigraDoc sample.

            // Create a renderer and prepare (=layout) the document.

            // For clarity we use point as unit of measure in this sample.
            // A4 is the standard letter size in Germany (21cm x 29.7cm).
            var a4Rect = new XRect(0, 0, A4Width, A4Height);

            // header
            var headerRect = GetHeaderRect();
            var headerContainer = gfx.BeginContainer(headerRect, headerRect, XGraphicsUnit.Point);
            headerRenderer.RenderPage(gfx, 1);
            gfx.EndContainer(headerContainer);
            // footer
            var footerRect = GetHeaderRect();
            var footerContainer = gfx.BeginContainer(footerRect, footerRect, XGraphicsUnit.Point);
            footerRenderer.RenderPage(gfx, 1);
            gfx.EndContainer(footerContainer);

            if (idx == 0)
            {
                // topbar
                var topRect = GetHeaderRect();
                var topContainer = gfx.BeginContainer(headerRect, headerRect, XGraphicsUnit.Point);
                topRenderer.RenderPage(gfx, 1);
                gfx.EndContainer(topContainer);
            }

            if (idx < sidePageCount)
            {
                // side bar
                var leftRect = GetLeftRect(idx);

                // Use BeginContainer / EndContainer for simplicity only. You can naturally use your own transformations.
                var leftContainer = gfx.BeginContainer(leftRect, leftRect, XGraphicsUnit.Point);

                // Draw page border for better visual representation.
                gfx.DrawRectangle(XPens.LightGray, leftRect);

                // Render the page. Note that page numbers start with 1.
                sideDocRenderer.RenderPage(gfx, idx + 1);

                // Note: The outline and the hyperlinks (table of contents) do not work in the produced PDF document.

                // Pop the previous graphical state.
                gfx.EndContainer(leftContainer);

            }


            // main content
            var rightRect = GetRightRect(idx);

            // Use BeginContainer / EndContainer for simplicity only. You can naturally use your own transformations.
            var container = gfx.BeginContainer(rightRect, rightRect, XGraphicsUnit.Point);

            // Draw page border for better visual representation.
            gfx.DrawRectangle(XPens.LightGray, a4Rect);

            // Render the page. Note that page numbers start with 1.
            docRenderer.RenderPage(gfx, idx + 1);

            // Note: The outline and the hyperlinks (table of contents) do not work in the produced PDF document.

            // Pop the previous graphical state.
            gfx.EndContainer(container);
        }

    }

    private XRect GetHeaderRect()
    {

        var rect = new XRect(0, 0, A4Width, A4Height);

        return rect;
    }


    private XRect GetLeftRect(int index)
    {
        //var rect = new XRect(0, 0, A4Width / 3 * 0.9, A4Height / 3 * 0.9)
        //{
        //    X = (index % 3) * A4Width / 3 + A4Width * 0.05 / 3,
        //    Y = (index / 3) * A4Height / 3 + A4Height * 0.05 / 3
        //};
        var width = XUnit.FromCentimeter(4).Point;
        var rect = new XRect(0, 0, width, A4Height);

        return rect;
    }

    private XRect GetRightRect(int index)
    {
        //var rect = new XRect(0, 0, A4Width / 3 * 0.9, A4Height / 3 * 0.9)
        //{
        //    X = (index % 3) * A4Width / 3 + A4Width * 0.05 / 3,
        //    Y = (index / 3) * A4Height / 3 + A4Height * 0.05 / 3
        //};
        var leftMargin = XUnit.FromCentimeter(4).Point;

        var rect = new XRect(leftMargin, 137, A4Width - leftMargin, A4Height);


        return rect;
    }

}


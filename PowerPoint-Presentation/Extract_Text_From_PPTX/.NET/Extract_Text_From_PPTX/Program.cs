using Syncfusion.Presentation;
using System.Text;

//Load the PowerPoint presentation 
using(IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Input.pptx")))
{    
    string extractedText = ExtractDocumentText(presentation);
    //Write the text collection to a text file
    System.IO.File.WriteAllText(Path.GetFullPath(@"Output/Output.txt"), extractedText);
}

static string ExtractDocumentText(IPresentation presentation)
{
    //Text collection to store the extracted text 
    StringBuilder textBuilder = new StringBuilder();
    // Extract text from all slides
    for (int i = 0; i < presentation.Slides.Count; i++)
    {
        ISlide slide = presentation.Slides[i];
        textBuilder.AppendLine($"--- Slide {i + 1} ---");

        // Extract text from all shapes in the slide
        ExtractText(slide.Shapes as IShapes, textBuilder);

        // Extract text from the slide notes body
        if (slide.NotesSlide?.NotesTextBody != null)
        {
            foreach (IParagraph paragraph in slide.NotesSlide.NotesTextBody.Paragraphs)
            {
                textBuilder.AppendLine(paragraph.Text);
            }
        }

        // Extract text from the slide notes shapes
        if (slide.NotesSlide?.Shapes != null)
        {
            ExtractText(slide.NotesSlide.Shapes as IShapes, textBuilder);
        }

        // Extract text from the layout slide shapes
        if (slide.LayoutSlide?.Shapes != null)
        {
            ExtractText(slide.LayoutSlide.Shapes as IShapes, textBuilder, true);

            // Extract text from the master slide shapes
            if (slide.LayoutSlide.MasterSlide?.Shapes != null)
            {
                ExtractText(slide.LayoutSlide.MasterSlide.Shapes as IShapes, textBuilder, true);
            }
        }
        textBuilder.AppendLine();
    }
    return textBuilder.ToString();
}

static void ExtractText(IShapes shapes, StringBuilder textBuilder, bool ignorePlaceHolder = false)
{
    foreach (IShape shape in shapes)
    {
        if (shape is ITable)
            ExtractTextInTable(shape, textBuilder);
        else if (shape is ISmartArt)
            ExtractTextInSmartArt(shape, textBuilder);
        else if (shape is IGroupShape)
            ExtractText((shape as IGroupShape).Shapes, textBuilder, ignorePlaceHolder);
        else
            ExtractTextInShape(shape, textBuilder, ignorePlaceHolder);
    }
}
static void ExtractTextInSmartArt(IShape shape, StringBuilder textBuilder)
{
    ISmartArt smartArt = shape as ISmartArt;
    if (smartArt == null)
        return;

    foreach (ISmartArtNode node in smartArt.Nodes)
    {
        ExtractTextInSmartArtNode(node, textBuilder);
    }
}
static void ExtractTextInShape(IShape shape, StringBuilder textBuilder, bool ignorePlaceHolder)
{
    if (shape.TextBody == null || (ignorePlaceHolder && (shape as ISlideItem).SlideItemType == SlideItemType.Placeholder))
        return;

    foreach (IParagraph paragraph in shape.TextBody.Paragraphs)
    {
        textBuilder.AppendLine(paragraph.Text);
    }
}
static void ExtractTextInTable(IShape shape, StringBuilder textBuilder)
{
    ITable table = shape as ITable;
    if (table == null)
        return;

    foreach (IRow row in table.Rows)
    {
        foreach (ICell cell in row.Cells)
        {
            textBuilder.AppendLine(cell.TextBody.Text);
        }
    }
}
static void ExtractTextInSmartArtNode(ISmartArtNode node, StringBuilder textBuilder)
{
    if (node.TextBody != null)
    {
        foreach (IParagraph paragraph in node.TextBody.Paragraphs)
        {
            textBuilder.AppendLine(paragraph.Text);
        }
    }

    // Recursively extract text from child nodes
    foreach (ISmartArtNode childNode in node.ChildNodes)
    {
        ExtractTextInSmartArtNode(childNode, textBuilder);
    }
}

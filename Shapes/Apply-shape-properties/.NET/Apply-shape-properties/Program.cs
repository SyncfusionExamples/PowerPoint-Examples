using Syncfusion.Presentation;

//Loads or opens a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide of the Presentation
ISlide slide = pptxDoc.Slides[0];
//Gets the first shape of the slide
IShape shape = slide.Shapes[0] as IShape;
//Sets the shape name
shape.ShapeName = "Shape1";
//Retrieves the line format of the shape
ILineFormat lineFormat = shape.LineFormat;
//Sets the dash style of the line format
lineFormat.DashStyle = LineDashStyle.DashDotDot;
//Sets the weight of the line format
lineFormat.Weight = 3;
//Sets the pattern fill type to the shape
shape.Fill.FillType = FillType.Pattern;
//Chooses the type of pattern
shape.Fill.PatternFill.Pattern = PatternFillType.DashedDownwardDiagonal;
//Sets the foreground color
shape.Fill.PatternFill.ForeColor = ColorObject.AliceBlue;
//Sets the background color
shape.Fill.PatternFill.BackColor = ColorObject.DarkSalmon;
//Saves the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
//Closes the Presentation
pptxDoc.Close();
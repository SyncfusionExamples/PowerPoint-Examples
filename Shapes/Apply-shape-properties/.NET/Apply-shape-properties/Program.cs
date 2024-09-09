using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide of the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the shape of the slide.
IShape shape = slide.Shapes[0] as IShape;
//Set the shape name.
shape.ShapeName = "Shape1";
//Retrieve the line format of the shape.
ILineFormat lineFormat = shape.LineFormat;
//Set the dash style of the line format.
lineFormat.DashStyle = LineDashStyle.DashDotDot;
//Set the weight of the line format.
lineFormat.Weight = 3;
//Set the pattern fill type to shape.
shape.Fill.FillType = FillType.Pattern;
//Choose the type of pattern. 
shape.Fill.PatternFill.Pattern = PatternFillType.DashedDownwardDiagonal;
//Set the fore color.
shape.Fill.PatternFill.ForeColor = ColorObject.AliceBlue;
//Set the back color.
shape.Fill.PatternFill.BackColor = ColorObject.DarkSalmon;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Gets the first slide
ISlide slide = pptxDoc.Slides[0];
//Get the chart in slide
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Change the chart type to 3D
chart.ChartType = OfficeChartType.Bar_Clustered_3D;
//Set the rotation
chart.Rotation = 80;
//Set the elevation
chart.Elevation = 90;
//Sets the shadow angle
chart.SideWall.Shadow.Angle = 60;
//Sets the back wall border weight
chart.BackWall.Border.LineWeight = OfficeChartLineWeight.Narrow;
//Set the right angle axes property of the chart
chart.RightAngleAxes = true;
//Set the auto scaling of chart
chart.AutoScaling = true;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
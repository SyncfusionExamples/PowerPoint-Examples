using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Opens the PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide
ISlide slide = pptxDoc.Slides[0];
//Gets the chart in slide
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Changes the chart type to 3D
chart.ChartType = OfficeChartType.Bar_Clustered_3D;
//Sets the rotation
chart.Rotation = 80;
//Sets the shadow angle
chart.SideWall.Shadow.Angle = 60;
//Sets the back wall border weight
chart.BackWall.Border.LineWeight = OfficeChartLineWeight.Narrow;
//Set the right angle axes property of the chart
chart.RightAngleAxes = true;
//Set the auto scaling of chart
chart.AutoScaling = true;
//Saves the Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();
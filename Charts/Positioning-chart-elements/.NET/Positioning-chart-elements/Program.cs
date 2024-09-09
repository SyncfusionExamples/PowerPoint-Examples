using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add chart to slide.
IPresentationChart chart = slide.Shapes.AddChart(100, 60, 800, 370);
//Set the data range of chart.
chart.DataRange = chart.ChartData[1, 1, 4, 3];
//Set data to the chart- RowIndex, columnIndex and data.
chart.ChartData.SetValue(1, 1, "Fruits");
chart.ChartData.SetValue(2, 1, "Apple");
chart.ChartData.SetValue(3, 1, "Orange");
chart.ChartData.SetValue(4, 1, "Mango");
//Set data to the chart- RowIndex, columnIndex and data.
chart.ChartData.SetValue(1, 2, "2010");
chart.ChartData.SetValue(2, 2, 330);
chart.ChartData.SetValue(3, 2, 490);
chart.ChartData.SetValue(4, 2, 700);
//Set data to the chart- RowIndex, columnIndex and data.
chart.ChartData.SetValue(1, 3, "2013");
chart.ChartData.SetValue(2, 3, 360);
chart.ChartData.SetValue(3, 3, 290);
chart.ChartData.SetValue(4, 3, 500);

chart.ChartType = OfficeChartType.Area;
//Edge: Specifies the width or Height to be interpreted as right or bottom of the chart element.
//Factor: Specifies the width or Height to be interpreted as the width or height of the chart element.
chart.PlotArea.Layout.LeftMode = LayoutModes.auto;
chart.PlotArea.Layout.TopMode = LayoutModes.factor;
//Value in points should not be negative value when LayoutMode is Edge.
//It can be a negative value, when the LayoutMode is Factor.
chart.ChartTitleArea.Layout.Left = 10;
chart.ChartTitleArea.Layout.Top = 100;
//Manually positions chart plot area.
chart.PlotArea.Layout.LayoutTarget = LayoutTargets.outer;
chart.PlotArea.Layout.LeftMode = LayoutModes.edge;
chart.PlotArea.Layout.TopMode = LayoutModes.edge;
//Manually positions chart legend.
chart.Legend.Layout.LeftMode = LayoutModes.factor;
chart.Legend.Layout.TopMode = LayoutModes.factor;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
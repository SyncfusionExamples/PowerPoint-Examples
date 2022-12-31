using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the chart from the first slide.
IPresentationChart chart = pptxDoc.Slides[0].Charts[0] as IPresentationChart;

//Set border settings.
chart.PlotArea.Border.AutoFormat = false;
//Set the auto line color.
chart.PlotArea.Border.IsAutoLineColor = false;
//Set the border line color.
chart.PlotArea.Border.LineColor = Color.Blue;
//Set the border line pattern.
chart.PlotArea.Border.LinePattern = OfficeChartLinePattern.DashDot;
//Set the border line weight.
chart.PlotArea.Border.LineWeight = OfficeChartLineWeight.Wide;
//Set the border transparency.
chart.PlotArea.Border.Transparency = 0.6;
//Set the plot area’s fill type.
chart.PlotArea.Fill.FillType = OfficeFillType.SolidColor;
//Set the plot area’s fill color.
chart.PlotArea.Fill.ForeColor = Color.LightPink;
//Set the plot area shadow presence.
chart.PlotArea.Shadow.ShadowInnerPresets = Office2007ChartPresetsInner.InsideDiagonalTopLeft;

//Set the legend position.
chart.Legend.Position = OfficeLegendPosition.Left;
//Set the layout inclusion.
chart.Legend.IncludeInLayout = true;
//Set the legend border format.
chart.Legend.FrameFormat.Border.AutoFormat = false;
//Set the legend border auto line color.
chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
//Set the border line color.
chart.Legend.FrameFormat.Border.LineColor = Color.Blue;
//Set the border line pattern.
chart.Legend.FrameFormat.Border.LinePattern = OfficeChartLinePattern.DashDot;
//Set the legend border line weight.
chart.Legend.FrameFormat.Border.LineWeight = OfficeChartLineWeight.Wide;
//Set the font weight for legend text.
chart.Legend.TextArea.Bold = true;
//Set the fore color to legend text.
chart.Legend.TextArea.Color = OfficeKnownColors.Bright_green;
//Set the font name for legend text.
chart.Legend.TextArea.FontName = "Times New Roman";
//Set the font size for legend text.
chart.Legend.TextArea.Size = 20;
//Set the legend text area' strike through.
chart.Legend.TextArea.Strikethrough = true;
//Modify the legend entry.
chart.Legend.LegendEntries[0].IsDeleted = true;
//Modify the legend layout height.
chart.Legend.Layout.Height = 100;
//Modify the legend layout height mode.
chart.Legend.Layout.HeightMode = LayoutModes.factor;
//Modify the legend layout left position.
chart.Legend.Layout.Left = 70;
//Modify the legend layout left mode.
chart.Legend.Layout.LeftMode = LayoutModes.factor;
//Modify the legend layout top position.
chart.Legend.Layout.Top = 70;
//Modify the legend layout top mode.
chart.Legend.Layout.TopMode = LayoutModes.factor;
//Modify the legend layout width.
chart.Legend.Layout.Width = 200;
//Modify the legend layout width mode.
chart.Legend.Layout.WidthMode = LayoutModes.factor;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
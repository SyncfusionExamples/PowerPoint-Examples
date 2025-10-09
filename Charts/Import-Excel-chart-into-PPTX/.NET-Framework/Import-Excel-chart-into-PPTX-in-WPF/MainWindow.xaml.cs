using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Import_Excel_chart_into_PPTX_in_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //Create the Excel file with column clustered chart by using XlsIO
            CreateExcelWithChart();
            //Import Excel chart in PowerPoint file by using Essential Presentation.
            ImportExcelChartInPresentation();
        }
        public static void CreateExcelWithChart()
        {
            //Create excel file and set application wise properties.
            ExcelEngine ExcelEngine = new ExcelEngine();
            IApplication Application = ExcelEngine.Excel;
            Application.DefaultVersion = ExcelVersion.Excel2013;
            IWorkbook WorkBook = Application.Workbooks.Create(1);
            IWorksheet Sheet = WorkBook.Worksheets[0];
            Sheet.IsRowColumnHeadersVisible = true;
            Sheet.IsGridLinesVisible = true;
            Sheet.SetColumnWidth(2, 20);
            //Give data for chart in sheet
            Sheet.Range["A1"].Text = "Company Name";
            Sheet.Range["B1"].Text = "1st Qatar";
            Sheet.Range["C1"].Text = "2nd Qatar";
            Sheet.Range["D1"].Text = "3rd Qatar";
            Sheet.Range["A2"].Text = "Microsoft";
            Sheet.Range["B2"].Number = 80000;
            Sheet.Range["C2"].Number = 80500;
            Sheet.Range["D2"].Number = 82000;
            Sheet.Range["A3"].Text = "Apple";
            Sheet.Range["B3"].Number = 76000;
            Sheet.Range["C3"].Number = 77000;
            Sheet.Range["D3"].Number = 80000;
            Sheet.Range["A4"].Text = "Dell";
            Sheet.Range["B4"].Number = 81000;
            Sheet.Range["C4"].Number = 79000;
            Sheet.Range["D4"].Number = 75000;
            Sheet.Range["A5"].Text = "Cisco";
            Sheet.Range["B5"].Number = 78000;
            Sheet.Range["C5"].Number = 83000;
            Sheet.Range["D5"].Number = 90000;
            Sheet.Range["A6"].Text = "Facebook";
            Sheet.Range["B6"].Number = 89000;
            Sheet.Range["C6"].Number = 92000;
            Sheet.Range["D6"].Number = 93000;
            //Add chart to the sheet
            IChartShape Chart = Sheet.Charts.Add();
            //Give data range for chart.
            Chart.DataRange = Sheet.Range["A1:D6"];
            // Give needed chart type
            Chart.ChartType = ExcelChartType.Column_Clustered;
            // Formmating chart area
            IChartFrameFormat ChartFormat = Chart.ChartArea;
            ChartFormat.Border.LinePattern = ExcelChartLinePattern.DarkGray;
            ChartFormat.Border.LineColor = System.Drawing.Color.Gray;
            ChartFormat.Border.LineWeight = ExcelChartLineWeight.Medium;
            ChartFormat.Fill.FillType = ExcelFillType.Gradient;
            ChartFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
            ChartFormat.Fill.GradientStyle = ExcelGradientStyle.Vertical;
            ChartFormat.Fill.BackColor = System.Drawing.Color.Goldenrod;
            ChartFormat.Fill.ForeColor = System.Drawing.Color.GreenYellow;
            // Give title to the chart
            Chart.ChartTitle = "Exam Report Chart By Clustered Chart View";
            Chart.PlotArea.Layout.LeftMode = Syncfusion.XlsIO.LayoutModes.edge;
            Chart.PlotArea.Layout.TopMode = Syncfusion.XlsIO.LayoutModes.edge;
            Chart.Legend.Layout.TopMode = Syncfusion.XlsIO.LayoutModes.edge;
            Chart.Legend.Layout.LeftMode = Syncfusion.XlsIO.LayoutModes.edge;
            Chart.PlotArea.Layout.Top = 20;
            Chart.PlotArea.Layout.Left = 2;
            Chart.Legend.Layout.Left = 2;
            Chart.Legend.Layout.Top = 20;
            Chart.HasDataTable = true;
            Chart.LeftColumn = 11;
            Chart.RightColumn = 21;
            Chart.TopRow = 10;
            Chart.BottomRow = 30;
            // Save the excel file.
            WorkBook.SaveAs("../../Charts.xlsx");
            WorkBook.Close();
            ExcelEngine.Dispose();
        }

        public static void ImportExcelChartInPresentation()
        {
            //Open the existing excel file.
            ExcelEngine ExcelEngine = new ExcelEngine();
            IApplication Application = ExcelEngine.Excel;
            Application.DefaultVersion = ExcelVersion.Excel2013;
            IWorkbook WorkBook = Application.Workbooks.Open("../../Charts.xlsx", ExcelOpenType.Automatic);
            IWorksheet sheet = WorkBook.Worksheets[0];
            // Create presentation by using Essential Presentation.
            IPresentation presentation = Presentation.Create();
            for (int i = 0; i < sheet.Charts.Count; i++)
            {
                //Get chart from excel sheet
                IChart ExcelChart = sheet.Charts[i] as IChart;
                //Add slide to the presentation document.
                ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
                //Add chart to the slide.
                IPresentationChart presentationchart = slide.Shapes.AddChart(30, 30, 500, 400);
                //Get the chart type of the excel chart and assign it in presentation chart.
                OfficeChartType ChartType = (OfficeChartType)ExcelChart.ChartType;
                //Get the data from excel chart.
                object[][] ChartData = new object[ExcelChart.DataRange.LastRow][];
                for (int a = 0; a < ExcelChart.DataRange.LastRow; a++)
                {
                    ChartData[a] = new object[ExcelChart.DataRange.LastColumn];
                }
                for (int b = 1, k = 0; b <= ExcelChart.DataRange.LastRow; b++)
                {
                    for (int j = 1, l = 0; j <= ExcelChart.DataRange.LastColumn; j++)
                    {
                        ChartData[k][l] = ExcelChart.DataRange[b, j].Value2;
                        l++;
                    }
                    k++;
                }
                //Set the excel chart data to the presentation chart.
                for (int j = 0; j < ChartData.Length; j++)
                {
                    for (int k = 0; k < ChartData[j].Length; k++)
                    {
                        presentationchart.ChartData.SetValue(j + 1, k + 1, ChartData[j][k]);
                    }
                }
                //Set the series values.
                for (int m = 0; m < ExcelChart.Series.Count; m++)
                {
                    int firstColum = ExcelChart.Series[m].Values.Column;
                    int firstRow = ExcelChart.Series[m].Values.Row;
                    int lastRow = ExcelChart.Series[m].Values.LastRow;
                    int lastColum = ExcelChart.Series[m].Values.LastColumn;
                    IOfficeChartSerie serie = presentationchart.Series.Add(ExcelChart.Series[m].Name.ToString());
                    serie.Values = presentationchart.ChartData[firstRow, firstColum, lastRow, lastColum];
                    serie.UsePrimaryAxis = ExcelChart.Series[m].UsePrimaryAxis;
                    serie.IsFiltered = ExcelChart.Series[m].IsFiltered;
                    serie.HasErrorBarsY = ExcelChart.Series[m].HasErrorBarsY;
                    serie.HasErrorBarsX = ExcelChart.Series[m].HasErrorBarsX;

                }
                presentationchart.ChartType = ChartType;
                for (int n = 0; n < ExcelChart.Series.Count; n++)
                {
                    presentationchart.Series[n].SerieType = (OfficeChartType)ExcelChart.Series[n].SerieType;               
                }
                //Copies the excel chart series format to the presentation chart series.
                for (int o = 0; o < ExcelChart.Series.Count; o++)
                {
                    CopiesExcelChartSeriesFormat(o, presentationchart.Series[o], ExcelChart);
                }
                int categoryFirstRow = ExcelChart.PrimaryCategoryAxis.CategoryLabels.Row;
                int categoryFirstColumn = ExcelChart.PrimaryCategoryAxis.CategoryLabels.Column;
                int categoryLastRow = ExcelChart.PrimaryCategoryAxis.CategoryLabels.LastRow;
                int categoryLastColumn = ExcelChart.PrimaryCategoryAxis.CategoryLabels.LastColumn;
                presentationchart.PrimaryCategoryAxis.CategoryLabels = presentationchart.ChartData[categoryFirstRow, categoryFirstColumn, categoryLastRow, categoryLastColumn];
                //Only use while using 3D-CHarts.    
                //presentationchart.AutoScaling = ExcelChart.AutoScaling;
                presentationchart.CategoryLabelLevel = (OfficeCategoriesLabelLevel)ExcelChart.CategoryLabelLevel;
                presentationchart.ChartTitle = ExcelChart.ChartTitle;
                //Copies excel chart area properties to the presentation chart area.
                CopiesExcelChartAreaProperties(presentationchart, ExcelChart);
                //Copies excel chart title area properties to the presentation chart title area.
                CopiesExcelChartTitleAreaProperties(presentationchart, ExcelChart);
                presentationchart.SizeWithWindow = ExcelChart.SizeWithWindow;
                //Only use it while using 3D-CHarts
                //presentationchart.WallsAndGridlines2D = ExcelChart.WallsAndGridlines2D;
                //Use the position of the Excel chart only when needed
                //presentationchart.XPos = ExcelChart.XPos;
                //presentationchart.YPos = ExcelChart.YPos;
                presentationchart.SeriesNameLevel = (OfficeSeriesNameLevel)ExcelChart.SeriesNameLevel;
                //Only use it while using 3D-CHarts
                //presentationchart.Rotation = ExcelChart.Rotation;
                //presentationchart.RightAngleAxes = ExcelChart.RightAngleAxes;
                presentationchart.PlotVisibleOnly = ExcelChart.PlotVisibleOnly;
                //Only use it while using 3D-Charts
                //presentationchart.Perspective = ExcelChart.Perspective;
                //presentationchart.HeightPercent = ExcelChart.HeightPercent;
                presentationchart.HasPlotArea = ExcelChart.HasPlotArea;
                presentationchart.HasLegend = ExcelChart.HasLegend;
                presentationchart.HasDataTable = ExcelChart.HasDataTable;
                //Only use it while using 3D-CHarts
                //presentationchart.GapDepth = ExcelChart.GapDepth;
                //presentationchart.Elevation = ExcelChart.Elevation;
                //presentationchart.DepthPercent = ExcelChart.DepthPercent;
            }
            presentation.Save("../../output.pptx");
        }
        public static void CopiesExcelChartSeriesFormat(int i, IOfficeChartSerie presentationChart, IChart ExcelChart)
        {
            presentationChart.SerieFormat.AreaProperties.BackgroundColor = ExcelChart.Series[i].SerieFormat.AreaProperties.BackgroundColor;
            presentationChart.SerieFormat.AreaProperties.BackgroundColorIndex = (OfficeKnownColors)ExcelChart.Series[i].SerieFormat.AreaProperties.BackgroundColorIndex;
            presentationChart.SerieFormat.AreaProperties.ForegroundColor = ExcelChart.Series[i].SerieFormat.AreaProperties.ForegroundColor;
            presentationChart.SerieFormat.AreaProperties.ForegroundColorIndex = (OfficeKnownColors)ExcelChart.Series[i].SerieFormat.AreaProperties.ForegroundColorIndex;
            presentationChart.SerieFormat.AreaProperties.Pattern = (OfficePattern)ExcelChart.Series[i].SerieFormat.AreaProperties.Pattern;
            presentationChart.SerieFormat.AreaProperties.SwapColorsOnNegative = ExcelChart.Series[i].SerieFormat.AreaProperties.SwapColorsOnNegative;
            presentationChart.SerieFormat.AreaProperties.UseAutomaticFormat = ExcelChart.Series[i].SerieFormat.AreaProperties.UseAutomaticFormat;
            presentationChart.SerieFormat.BarShapeBase = (OfficeBaseFormat)ExcelChart.Series[i].SerieFormat.BarShapeBase;
            presentationChart.SerieFormat.BarShapeTop = (OfficeTopFormat)ExcelChart.Series[i].SerieFormat.BarShapeTop;
            if (ExcelChart.ChartType == ExcelChartType.Bubble || ExcelChart.ChartType == ExcelChartType.Bubble_3D)
            {
                presentationChart.SerieFormat.CommonSerieOptions.BubbleScale = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.BubbleScale;
                presentationChart.SerieFormat.CommonSerieOptions.ShowNegativeBubbles = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.ShowNegativeBubbles;
                presentationChart.SerieFormat.Is3DBubbles = ExcelChart.Series[i].SerieFormat.Is3DBubbles;
                presentationChart.SerieFormat.CommonSerieOptions.SizeRepresents = (ChartBubbleSize)ExcelChart.Series[i].SerieFormat.CommonSerieOptions.SizeRepresents;
            }
            if (ExcelChart.ChartType == ExcelChartType.Doughnut)
            {
                presentationChart.SerieFormat.CommonSerieOptions.DoughnutHoleSize = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.DoughnutHoleSize;
                presentationChart.SerieFormat.Percent = ExcelChart.Series[i].SerieFormat.Percent;
                presentationChart.SerieFormat.CommonSerieOptions.FirstSliceAngle = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.FirstSliceAngle;
            }
            if (ExcelChart.ChartType != ExcelChartType.Pie || ExcelChart.ChartType != ExcelChartType.Pie_Bar || ExcelChart.ChartType != ExcelChartType.PieOfPie)
            {
                //presentationChart.SerieFormat.CommonSerieOptions.GapWidth = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.GapWidth;
                //presentationChart.SerieFormat.CommonSerieOptions.Overlap = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.Overlap;
            }
            if (ExcelChart.ChartType == ExcelChartType.Radar)
            {
                presentationChart.SerieFormat.CommonSerieOptions.HasRadarAxisLabels = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.HasRadarAxisLabels;
            }
            presentationChart.SerieFormat.CommonSerieOptions.IsVaryColor = ExcelChart.Series[i].SerieFormat.CommonSerieOptions.IsVaryColor;

            if (ExcelChart.ChartType == ExcelChartType.Line_Markers || ExcelChart.ChartType == ExcelChartType.Line_Markers_Stacked || ExcelChart.ChartType == ExcelChartType.Line_Markers_Stacked_100)
            {
                presentationChart.SerieFormat.IsAutoMarker = ExcelChart.Series[i].SerieFormat.IsAutoMarker;
                presentationChart.SerieFormat.MarkerBackgroundColor = ExcelChart.Series[i].SerieFormat.MarkerBackgroundColor;
                presentationChart.SerieFormat.MarkerBackgroundColorIndex = (OfficeKnownColors)ExcelChart.Series[i].SerieFormat.MarkerBackgroundColorIndex;
                presentationChart.SerieFormat.MarkerForegroundColor = ExcelChart.Series[i].SerieFormat.MarkerForegroundColor;
                presentationChart.SerieFormat.MarkerForegroundColorIndex = (OfficeKnownColors)ExcelChart.Series[i].SerieFormat.MarkerForegroundColorIndex;
                presentationChart.SerieFormat.MarkerSize = ExcelChart.Series[i].SerieFormat.MarkerSize;
                presentationChart.SerieFormat.MarkerStyle = (OfficeChartMarkerType)ExcelChart.Series[i].SerieFormat.MarkerStyle;
            }

            presentationChart.SerieFormat.Shadow.Angle = ExcelChart.Series[i].SerieFormat.Shadow.Angle;
            presentationChart.SerieFormat.Shadow.Blur = ExcelChart.Series[i].SerieFormat.Shadow.Blur;
            presentationChart.SerieFormat.Shadow.Distance = ExcelChart.Series[i].SerieFormat.Shadow.Distance;
            presentationChart.SerieFormat.Shadow.HasCustomShadowStyle = ExcelChart.Series[i].SerieFormat.Shadow.HasCustomShadowStyle;
            presentationChart.SerieFormat.Shadow.ShadowColor = ExcelChart.Series[i].SerieFormat.Shadow.ShadowColor;
            presentationChart.SerieFormat.Shadow.ShadowInnerPresets = (Office2007ChartPresetsInner)ExcelChart.Series[i].SerieFormat.Shadow.ShadowInnerPresets;
            presentationChart.SerieFormat.Shadow.ShadowOuterPresets = (Office2007ChartPresetsOuter)ExcelChart.Series[i].SerieFormat.Shadow.ShadowOuterPresets;
            presentationChart.SerieFormat.Shadow.ShadowPerspectivePresets = (Office2007ChartPresetsPerspective)ExcelChart.Series[i].SerieFormat.Shadow.ShadowPrespectivePresets;
            if (ExcelChart.Series[i].SerieFormat.Shadow.Size > 0)
            {
                presentationChart.SerieFormat.Shadow.Size = ExcelChart.Series[i].SerieFormat.Shadow.Size;
            }
            presentationChart.SerieFormat.Shadow.Transparency = ExcelChart.Series[i].SerieFormat.Shadow.Transparency;
            presentationChart.SerieFormat.ThreeD.BevelBottom = (Office2007ChartBevelProperties)ExcelChart.Series[i].SerieFormat.ThreeD.BevelBottom;
            presentationChart.SerieFormat.ThreeD.BevelBottomHeight = ExcelChart.Series[i].SerieFormat.ThreeD.BevelBottomHeight;
            presentationChart.SerieFormat.ThreeD.BevelBottomWidth = ExcelChart.Series[i].SerieFormat.ThreeD.BevelBottomWidth;
            presentationChart.SerieFormat.ThreeD.BevelTop = (Office2007ChartBevelProperties)ExcelChart.Series[i].SerieFormat.ThreeD.BevelTop;
            presentationChart.SerieFormat.ThreeD.BevelTopHeight = ExcelChart.Series[i].SerieFormat.ThreeD.BevelTopHeight;
            presentationChart.SerieFormat.ThreeD.BevelTopWidth = ExcelChart.Series[i].SerieFormat.ThreeD.BevelTopWidth;
            presentationChart.SerieFormat.ThreeD.Lighting = (Office2007ChartLightingProperties)ExcelChart.Series[i].SerieFormat.ThreeD.Lighting;
            presentationChart.SerieFormat.ThreeD.Material = (Office2007ChartMaterialProperties)ExcelChart.Series[i].SerieFormat.ThreeD.Material;
        }
        public static void CopiesExcelChartAreaProperties(IPresentationChart presentationChart, IChart ExcelChart)
        {
            presentationChart.ChartArea.Border.AutoFormat = ExcelChart.ChartArea.Border.AutoFormat;
            presentationChart.ChartArea.Border.ColorIndex = (OfficeKnownColors)ExcelChart.ChartArea.Border.ColorIndex;
            presentationChart.ChartArea.Border.IsAutoLineColor = ExcelChart.ChartArea.Border.IsAutoLineColor;
            presentationChart.ChartArea.Border.LineColor = ExcelChart.ChartArea.Border.LineColor;
            presentationChart.ChartArea.Border.LinePattern = (OfficeChartLinePattern)ExcelChart.ChartArea.Border.LinePattern;
            presentationChart.ChartArea.Border.LineWeight = (OfficeChartLineWeight)ExcelChart.ChartArea.Border.LineWeight;
            presentationChart.ChartArea.Border.Transparency = ExcelChart.ChartArea.Border.Transparency;
            presentationChart.ChartArea.Fill.FillType = (OfficeFillType)ExcelFillType.Gradient;
            presentationChart.ChartArea.Fill.GradientColorType = (OfficeGradientColor)ExcelGradientColor.TwoColor;
            presentationChart.ChartArea.Fill.GradientStyle = (OfficeGradientStyle)ExcelGradientStyle.Vertical;
            presentationChart.ChartArea.Fill.BackColor = ExcelChart.ChartArea.Fill.BackColor;
            presentationChart.ChartArea.Fill.ForeColor = ExcelChart.ChartArea.Fill.ForeColor;
        }
        public static void CopiesExcelChartTitleAreaProperties(IPresentationChart presentationChart, IChart ExcelChart)
        {
            presentationChart.ChartTitleArea.Bold = ExcelChart.ChartTitleArea.Bold;
            presentationChart.ChartTitleArea.Color = (OfficeKnownColors)ExcelChart.ChartTitleArea.Color;
            presentationChart.ChartTitleArea.FontName = ExcelChart.ChartTitleArea.FontName;
            presentationChart.ChartTitleArea.FrameFormat.Border.AutoFormat = ExcelChart.ChartTitleArea.FrameFormat.Border.AutoFormat;
            presentationChart.ChartTitleArea.FrameFormat.Border.ColorIndex = (OfficeKnownColors)ExcelChart.ChartTitleArea.FrameFormat.Border.ColorIndex;
            presentationChart.ChartTitleArea.FrameFormat.Border.IsAutoLineColor = ExcelChart.ChartTitleArea.FrameFormat.Border.IsAutoLineColor;
            presentationChart.ChartTitleArea.FrameFormat.Border.LineColor = ExcelChart.ChartTitleArea.FrameFormat.Border.LineColor;
            presentationChart.ChartTitleArea.FrameFormat.Border.LinePattern = (OfficeChartLinePattern)ExcelChart.ChartTitleArea.FrameFormat.Border.LinePattern;
            presentationChart.ChartTitleArea.FrameFormat.Border.LineWeight = (OfficeChartLineWeight)ExcelChart.ChartTitleArea.FrameFormat.Border.LineWeight;
            presentationChart.ChartTitleArea.FrameFormat.Border.Transparency = ExcelChart.ChartTitleArea.FrameFormat.Border.Transparency;
            presentationChart.ChartTitleArea.FrameFormat.IsBorderCornersRound = ExcelChart.ChartTitleArea.FrameFormat.IsBorderCornersRound;
        }
    }
}

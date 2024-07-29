using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Create_PowerPoint_presentation
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void CreatePresentation(object sender, RoutedEventArgs e)
        {
            //Create a new instance of the PowerPoint Presentation file.
            using IPresentation pptxDoc = Presentation.Create();
            //Add a new slide to the file and apply background color.
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);
            //Specify the fill type and fill color for the slide background.
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);
            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide.
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph("Company History").HorizontalAlignment = HorizontalAlignmentType.Center;
            //Add description content to the slide by adding a new TextBox.
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = "IMN Solutions PVT LTD is the software company, established in 1987, by George Milton. The company has been listed as the trusted partner for many high-profile organizations since 1988 and got awards for quality products from reputed organizations.";
            //Add bullet points to the slide.
            IShape bulletPointsShape = slide.AddTextBox(53.22, 270, 437.90, 116.32);
            //Add a paragraph for a bullet point.
            IParagraph firstPara = bulletPointsShape.TextBody.AddParagraph("The company acquired the MCY corporation for 20 billion dollars and became the top revenue maker for the year 2015.");
            //Format how the bullets should be displayed.
            firstPara.ListFormat.Type = ListType.Bulleted;
            firstPara.LeftIndent = 35;
            firstPara.FirstLineIndent = -35;
            //Add another paragraph for the next bullet point.
            IParagraph secondPara = bulletPointsShape.TextBody.AddParagraph("The company is participating in top open source projects in automation industry.");
            //Format how the bullets should be displayed.
            secondPara.ListFormat.Type = ListType.Bulleted;
            secondPara.LeftIndent = 35;
            secondPara.FirstLineIndent = -35;
            //Get a picture as stream.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            Stream pictureStream = assembly.GetManifestResourceStream("Create_PowerPoint_presentation.Assets.Image.jpg");
            //Add the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            //Add an auto-shape to the slide.
            IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);
            //Format the auto-shape color by setting the fill type and text.
            stampShape.Fill.FillType = FillType.None;
            stampShape.TextBody.AddParagraph("IMN").HorizontalAlignment = HorizontalAlignmentType.Center;
            //Save the PowerPoint Presentation document to stream.
            using MemoryStream pptx = new();
            pptxDoc.Save(pptx);
            //Saves and launch the file.
            SaveHelper.SaveAndLaunch("Sample.pptx", pptx);
        }
    }
}

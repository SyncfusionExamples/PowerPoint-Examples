using System;
using System.IO;
using System.Reflection;
using Microsoft.Maui.Controls;
using CreatePowerPoint.Services;
using Syncfusion.Presentation;

namespace CreatePowerPoint
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void CreatePresentation(object sender, EventArgs e)
        {
            //Creates a new instance of the PowerPoint Presentation file.
            using IPresentation pptxDoc = Presentation.Create();
            //Adds a new slide to the file and apply background color.
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);
            //Specifies the fill type and fill color for the slide background.
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);

            //Adds title content to the slide by accessing the title placeholder of the TitleOnly layout-slide.
            Syncfusion.Presentation.IShape titleShape = slide.Shapes[0] as Syncfusion.Presentation.IShape;
            titleShape.TextBody.AddParagraph("Company History").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Adds description content to the slide by adding a new TextBox.
            Syncfusion.Presentation.IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = "IMN Solutions PVT LTD is the software company, established in 1987, by George Milton. The company has been listed as the trusted partner for many high-profile organizations since 1988 and got awards for quality products from reputed organizations.";

            //Adds bullet points to the slide.
            Syncfusion.Presentation.IShape bulletPointsShape = slide.AddTextBox(53.22, 270, 437.90, 116.32);
            //Adds a paragraph for a bullet point.
            IParagraph firstPara = bulletPointsShape.TextBody.AddParagraph("The company acquired the MCY corporation for 20 billion dollars and became the top revenue maker for the year 2015.");
            //Formats how the bullets should be displayed.
            firstPara.ListFormat.Type = ListType.Bulleted;
            firstPara.LeftIndent = 35;
            firstPara.FirstLineIndent = -35;
            //Adds another paragraph for the next bullet point.
            IParagraph secondPara = bulletPointsShape.TextBody.AddParagraph("The company is participating in top open source projects in automation industry.");
            //Formats how the bullets should be displayed.
            secondPara.ListFormat.Type = ListType.Bulleted;
            secondPara.LeftIndent = 35;
            secondPara.FirstLineIndent = -35;

            Assembly assembly = typeof(MainPage).GetTypeInfo().Assembly;
            string resourcePath = "CreatePowerPoint.Resources.Presentation.Image.png";
            //Gets a picture as stream.
            Stream pictureStream = assembly.GetManifestResourceStream(resourcePath);
            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Adds an auto-shape to the slide.
            Syncfusion.Presentation.IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);
            //Formats the auto-shape color by setting the fill type and text.
            stampShape.Fill.FillType = FillType.None;
            stampShape.TextBody.AddParagraph("IMN").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Saves the presentation to the memory stream.
            using MemoryStream stream = new();
            pptxDoc.Save(stream);
            stream.Position = 0;
            //Saves the memory stream as file.
			SaveService saveService = new();
            saveService.SaveAndView("Sample.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", stream);
        }
    }

    
}

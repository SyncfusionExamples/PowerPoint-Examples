using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System.Reflection;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.Storage.Pickers;
using Syncfusion.Presentation;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Create_PowerPoint_presentation
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private void CreatePresentation(object sender, RoutedEventArgs e)
        {
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Create a new instance of the PowerPoint Presentation file.
            using (IPresentation pptxDoc = Presentation.Create())
            {
                //Add a new slide to file and apply background color.
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
                Stream pictureStream = assembly.GetManifestResourceStream("Create_PowerPoint_presentation.Assets.Image.jpg");
                //Add the picture to a slide by specifying its size and position.
                slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
                //Add an auto-shape to the slide.
                IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);
                //Format the auto-shape color by setting the fill type and text.
                stampShape.Fill.FillType = FillType.None;
                stampShape.TextBody.AddParagraph("IMN").HorizontalAlignment = HorizontalAlignmentType.Center;
                //Save the Presentation file to MemoryStream.
                using (MemoryStream stream = new MemoryStream())
                {
                    pptxDoc.Save(stream);
                    //Save the stream as a Presentation file in the local machine.
                    Save(stream);
                }
            }

            async void Save(MemoryStream stream)
            {
                StorageFile stFile;
                if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
                {
                    FileSavePicker savePicker = new FileSavePicker();
                    savePicker.DefaultFileExtension = ".pptx";
                    savePicker.SuggestedFileName = "Sample";
                    savePicker.FileTypeChoices.Add("PowerPoint Files", new List<string>() { ".pptx" });
                    stFile = await savePicker.PickSaveFileAsync();
                }
                else
                {
                    StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                    stFile = await local.CreateFileAsync("Sample", CreationCollisionOption.ReplaceExisting);
                }
                if (stFile != null)
                {
                    using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                    {
                        //Write compressed data from memory to file.
                        using (Stream outstream = zipStream.AsStreamForWrite())
                        {
                            byte[] buffer = stream.ToArray();
                            outstream.Write(buffer, 0, buffer.Length);
                            outstream.Flush();
                        }
                    }
                }
                //Launch the saved PowerPoint file.
                await Windows.System.Launcher.LaunchFileAsync(stFile);
            }
        }
    }
}

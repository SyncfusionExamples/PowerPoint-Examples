using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Syncfusion.PresentationRenderer;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_chart_to_image
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
        private async void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //"App" is the class of Portable project.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Open an existing PowerPoint presentation
            IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert_chart_to_image.Assets.Template.pptx"));
            //Initialize the PresentationRenderer
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Gets the first instance of chart from slide
            IPresentationChart chart = pptxDoc.Slides[0].Charts[0];
            // Converts the chart to image.
            //Creates a stream instance to store the image
            MemoryStream stream = new MemoryStream();
            pptxDoc.PresentationRenderer.ConvertToImage(chart, stream);
            //Closes the presentation
            pptxDoc.Close();
            //Save the memory stream as file.
            Save(stream as MemoryStream, "ChartToImage.jpeg");
        }
        /// <summary>
        /// Save the image.
        /// </summary>
        async void Save(MemoryStream streams, string filename)
        {
            streams.Position = 0;
            StorageFile stFile;
            if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".jpeg";
                savePicker.SuggestedFileName = filename;
                savePicker.FileTypeChoices.Add("Image", new List<string>() { ".jpeg" });
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync(filename, CreationCollisionOption.ReplaceExisting);
            }
            if (stFile != null)
            {
                using (Windows.Storage.Streams.IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    //Write compressed data from memory to file.
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                        byte[] buffer = streams.ToArray();
                        outstream.Write(buffer, 0, buffer.Length);
                        outstream.Flush();
                    }
                }
            }
            //Launch the saved image file.
            await Windows.System.Launcher.LaunchFileAsync(stFile);
        }
    }
}

using Syncfusion.Presentation;
using System;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Change_resolution_of_converted_image
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
            //Load the presentation file using open picker
            FileOpenPicker openPicker = new FileOpenPicker();
            openPicker.FileTypeFilter.Add(".pptx");
            StorageFile inputFile = await openPicker.PickSingleFileAsync();
            IPresentation pptxDoc = await Presentation.OpenAsync(inputFile);
            //Initialize the ‘ChartToImageConverter’ instance to convert the charts in the slides
            pptxDoc.ChartToImageConverter = new ChartToImageConverter();
            //Pick the folder to save the converted images.
            FolderPicker folderPicker = new FolderPicker();
            folderPicker.ViewMode = PickerViewMode.Thumbnail;
            folderPicker.FileTypeFilter.Add("*");
            StorageFolder storageFolder = await folderPicker.PickSingleFolderAsync();
            StorageFile imageFile = await storageFolder.CreateFileAsync("Slide1.jpg", CreationCollisionOption.ReplaceExisting);
            //Get the stream of the created image file.
            Stream imageStream = await imageFile.OpenStreamForWriteAsync();
            //Creates a new instance for the rendering options to customize the image resolution.
            RenderingOptions renderingOptions = new RenderingOptions();
            //Sets the horizontal scaling value for the converted image. The default value is 1.
            renderingOptions.ScaleX = 10F;
            //Sets the vertical scaling value for the converted image. The default value is 1.
            renderingOptions.ScaleY = 10F;
            ISlide slide = pptxDoc.Slides[0];
            //Convert the slide to image with specified resolution.
            await slide.SaveAsImageAsync(imageStream, renderingOptions);
            //Closes the presentation instance
            pptxDoc.Close();
        }
    }
}

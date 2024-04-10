using Syncfusion.Presentation;
using System;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using Windows.UI.Xaml.Documents;
using System.Threading;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_with_UWP_cancellation_token
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
            FolderPicker folderPicker = new FolderPicker
            {
                ViewMode = PickerViewMode.Thumbnail
            };
            folderPicker.FileTypeFilter.Add("*");
            StorageFolder storageFolder = await folderPicker.PickSingleFolderAsync();
            //Create a cancellation token to cancel the image rendering instantly.
            CancellationTokenSource cancellationToken = new CancellationTokenSource();
            //Convert the slide to image.
            int slideNumber = 1;
            foreach (ISlide slide in pptxDoc.Slides)
            {
                StorageFile imageFile = await storageFolder.CreateFileAsync("Slide" + slideNumber++ + ".jpg", CreationCollisionOption.ReplaceExisting);
                await slide.SaveAsImageAsync(imageFile, cancellationToken.Token);
            }
        }
    }
}

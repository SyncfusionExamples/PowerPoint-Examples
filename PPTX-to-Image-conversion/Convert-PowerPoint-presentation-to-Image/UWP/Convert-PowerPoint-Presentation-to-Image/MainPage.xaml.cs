using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Syncfusion.Presentation;
using System.Reflection;
using Windows.Data.Pdf;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Popups;
using Windows.Graphics.Imaging;
using Syncfusion.OfficeChartToImageConverter;
using Windows.UI.Xaml.Documents;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_PowerPoint_Presentation_to_Image
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
            //Load the presentation file using open picker.
            FileOpenPicker openPicker = new FileOpenPicker();
            openPicker.FileTypeFilter.Add(".pptx");
            StorageFile inputFile = await openPicker.PickSingleFileAsync();
            using (IPresentation pptxDoc = await Presentation.OpenAsync(inputFile))
            {
                //Initialize the ‘ChartToImageConverter’ instance to convert the charts in the slides.
                pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                //Pick the folder to save the converted images.
                FolderPicker folderPicker = new FolderPicker
                {
                    ViewMode = PickerViewMode.Thumbnail
                };
                folderPicker.FileTypeFilter.Add("*");
                StorageFolder storageFolder = await folderPicker.PickSingleFolderAsync();
                StorageFile imageFile = await storageFolder.CreateFileAsync("PPTXtoImage.jpg", CreationCollisionOption.ReplaceExisting);
                ISlide slide = pptxDoc.Slides[0];
                //Convert the PPTX to image.
                await slide.SaveAsImageAsync(imageFile);
            }
        }
    }
}

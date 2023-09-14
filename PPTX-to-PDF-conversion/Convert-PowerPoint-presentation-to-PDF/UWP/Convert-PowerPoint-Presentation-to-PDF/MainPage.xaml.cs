using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
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
using Syncfusion.Pdf;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_PowerPoint_Presentation_to_PDF
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
            IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert_PowerPoint_Presentation_to_PDF.Assets.Input.pptx"));

            //Convert the PowerPoint document to PDF document.
            using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
            {
                // Create a MemoryStream to hold the PDF data.
                MemoryStream pdfStream = new MemoryStream();

                // Save the converted PDF document to the MemoryStream.
                pdfDocument.Save(pdfStream);

                // Initializes FileSavePicker
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.SuggestedStartLocation = PickerLocationId.Desktop;
                savePicker.SuggestedFileName = "Sample";
                savePicker.FileTypeChoices.Add("PDF Files", new List<string>() { ".pdf" });

                // Creates a storage file from FileSavePicker
                StorageFile storageFile = await savePicker.PickSaveFileAsync();

                if (storageFile != null)
                {
                    // Save the PDF data to the selected file.
                    using (var storageStream = await storageFile.OpenStreamForWriteAsync())
                    {
                        await pdfStream.CopyToAsync(storageStream);
                    }
                }
            }         
        }
    }
}

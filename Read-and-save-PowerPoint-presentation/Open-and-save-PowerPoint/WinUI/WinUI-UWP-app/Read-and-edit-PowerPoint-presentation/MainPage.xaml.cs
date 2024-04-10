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

namespace Read_and_edit_PowerPoint_presentation
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
            //"App" is the class of Portable project.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Open an existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Read_and_edit_PowerPoint_presentation.Assets.Template.pptx")))
            {
                //Get the first slide from the PowerPoint presentation.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the first shape of the slide.
                IShape shape = slide.Shapes[0] as IShape;
                //Modify the text of the shape.
                if (shape.TextBody.Text == "Company History")
                    shape.TextBody.Text = "Company Profile";
                //Save the Presentation files to MemoryStream.
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

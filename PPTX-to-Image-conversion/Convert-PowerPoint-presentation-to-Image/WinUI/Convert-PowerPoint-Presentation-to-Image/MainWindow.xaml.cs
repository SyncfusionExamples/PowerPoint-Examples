using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
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

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Convert_PowerPoint_Presentation_to_Image
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

        private void ConvertPPTXtoImage(object sender, RoutedEventArgs e)
        {
            //Loading an existing PowerPoint document.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Open the existing PowerPoint presentation with loaded stream.
            using (IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert_PowerPoint_Presentation_to_Image.Assets.Input.pptx")))
            {
                //Initialize the PresentationRenderer.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Converts the first slide into image.
                Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                //Reset the stream position.
                stream.Position = 0;
                //Save the stream as a image file in the local machine.
                SaveHelper.SaveAndLaunch("PPTXtoImage.Jpeg", stream as MemoryStream);
            }
        }
    }
}

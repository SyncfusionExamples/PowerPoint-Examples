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

namespace Read_and_edit_PowerPoint_presentation
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

        private void CreateDocument(object sender, RoutedEventArgs e)
        {
            //Get a picture as stream.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            Stream stream = assembly.GetManifestResourceStream("Read_and_edit_PowerPoint_presentation.Assets.Template.pptx");
            //Opens an existing PowerPoint presentation.
            using IPresentation pptxDoc = Presentation.Open(stream);
            //Get the first slide from the PowerPoint presentation.
            ISlide slide = pptxDoc.Slides[0];
            //Get the first shape of the slide.
            IShape shape = slide.Shapes[0] as IShape;
            //Modify the text of the shape.
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Save the PowerPoint Presentation document to stream.
            using MemoryStream pptx = new();
            pptxDoc.Save(pptx);
            //Saves and launch the file.
            SaveHelper.SaveAndLaunch("Sample.pptx", pptx);
        }
    }
}

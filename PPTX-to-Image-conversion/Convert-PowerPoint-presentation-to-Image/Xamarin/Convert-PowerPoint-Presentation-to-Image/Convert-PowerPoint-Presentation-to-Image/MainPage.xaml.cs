
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Convert_PowerPoint_Presentation_to_Image
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            //Loading an existing PowerPoint document.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Open the existing PowerPoint presentation with loaded stream.
            using (IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert-PowerPoint-Presentation-to-Image.Assets.Input.pptx")))
            {
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert PowerPoint slide to image as stream.
                using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                {
                    //Reset the stream position.
                    stream.Position = 0;
                    //Create the MemoryStream to save the converted image.
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        stream.CopyTo(imageStream);
                        //Save the stream as a file in the device and invoke it for viewing.
                        Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("PPTXtoImage.Jpeg", "application/jpeg", imageStream);
                    }
                }                   
            }
        }
    }
}

using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Read_and_edit_PowerPoint_presentation
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            //"App" is the class of Portable project.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            //Open an existing PowerPoint presentation
            IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Read-and-edit-PowerPoint-presentation.Assets.Template.pptx"));
            //Gets the first slide from the PowerPoint presentation
            ISlide slide = pptxDoc.Slides[0];
            //Gets the first shape of the slide
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Create new memory stream to save Presentation.
            MemoryStream stream = new MemoryStream();
            //Save Presentation in stream format.
            pptxDoc.Save(stream);
            //Close the presentation
            pptxDoc.Close();
            stream.Position = 0;
            //Save the stream as a file in the device and invoke it for viewing
            Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("Output.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", stream);
        }
    }
}

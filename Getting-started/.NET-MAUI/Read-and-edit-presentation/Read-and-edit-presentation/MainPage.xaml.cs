using Syncfusion.Presentation;
using System.Reflection;
using ReadPowerPoint.Services;

namespace Read_and_edit_presentation
{
    public partial class MainPage : ContentPage
    {
        int count = 0;

        public MainPage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Read and edit a PowerPoint file.
        /// </summary>
        private void ReadPowerPoint(object sender, EventArgs e)
        {
            Assembly assembly = typeof(MainPage).GetTypeInfo().Assembly;
            //Opens an existing PowerPoint presentation.
            using IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Read_and_edit_presentation.Resources.Presentation.Sample.pptx"));

            //Gets the first slide from the PowerPoint presentation.
            ISlide slide = pptxDoc.Slides[0];

            //Gets the first shape of the slide.
            Syncfusion.Presentation.IShape shape = slide.Shapes[0] as Syncfusion.Presentation.IShape;

            //Modifies the text of the shape.
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Saves the presentation to the memory stream.
            using MemoryStream stream = new();
            pptxDoc.Save(stream);
            stream.Position = 0;
            //Saves the memory stream as file.
            SaveService saveService = new();
            saveService.SaveAndView("Output.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", stream);
        }
    }

}

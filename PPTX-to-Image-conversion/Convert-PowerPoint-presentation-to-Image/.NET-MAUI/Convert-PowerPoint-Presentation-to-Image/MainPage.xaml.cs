using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.Reflection;
using System.Drawing;
using System.IO;
using Convert_PowerPoint_Presentation_to_Image.SaveServices;

namespace Convert_PowerPoint_Presentation_to_Image;

public partial class MainPage : ContentPage
{
	int count = 0;

	public MainPage()
	{
		InitializeComponent();
	}

	private void ConvertPPTXtoImage(object sender, EventArgs e)
	{
        //Loading an existing PowerPoint presentation
        Assembly assembly = typeof(App).GetTypeInfo().Assembly;
		//Open the existing PowerPoint presentation with loaded stream.
		using (IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert_PowerPoint_Presentation_to_Image.Assets.Input.pptx")))
		{
            //Initialize the PresentationRenderer.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Converts the first slide into image.
            Stream stream= pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
            //Reset the stream position.
            stream.Position = 0;
            //save and Launch the image file.
            SaveService saveService = new();
            saveService.SaveAndView("PPTXtoImage.Jpeg", "application/jpeg", stream as MemoryStream);
        }
    }
}


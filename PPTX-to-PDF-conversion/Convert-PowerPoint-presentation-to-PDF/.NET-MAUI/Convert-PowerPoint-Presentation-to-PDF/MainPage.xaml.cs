using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.Reflection;
using Convert_PowerPoint_Presentation_to_PDF.SaveServices;

namespace Convert_PowerPoint_Presentation_to_PDF;

public partial class MainPage : ContentPage
{
	int count = 0;

	public MainPage()
	{
		InitializeComponent();
	}

    private void ConvertPPTXtoPDF(object sender, EventArgs e)
    {
        //Loading an existing PowerPoint presentation
        Assembly assembly = typeof(App).GetTypeInfo().Assembly;
        //Open the existing PowerPoint presentation with loaded stream.
        using (IPresentation pptxDoc = Presentation.Open(assembly.GetManifestResourceStream("Convert_PowerPoint_Presentation_to_PDF.Assets.Input.pptx")))
        {
            //Convert the PowerPoint document to PDF document.
            using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
            {
                //Save the converted PDF document to MemoryStream.
                MemoryStream pdfStream = new MemoryStream();
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;
                //save and Launch the PDF document
                SaveService saveService = new();
                saveService.SaveAndView("Sample.pdf", "application/pdf", pdfStream);
            }
        }
    }
}


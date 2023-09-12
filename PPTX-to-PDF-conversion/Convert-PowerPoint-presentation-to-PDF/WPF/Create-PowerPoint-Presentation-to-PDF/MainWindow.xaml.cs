using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System.IO;
using System.ComponentModel;
using System.Windows;


namespace Create_PowerPoint_Presentation_to_PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            //Opens a PowerPoint Presentation
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"../../Data/Input.pptx")))
            {
                //Converts the PowerPoint Presentation into PDF document
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Saves the PDF document
                    pdfDocument.Save(Path.GetFullPath(@"../../Sample.pdf"));
                }                   
            }               
            //Launch the PDF document
            if (System.Windows.MessageBox.Show("Do you want to view the PDF document?", "Pdf document created",
  MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../Sample.pdf")) { UseShellExecute = true };
                process.Start();
            }
        }
    }
}

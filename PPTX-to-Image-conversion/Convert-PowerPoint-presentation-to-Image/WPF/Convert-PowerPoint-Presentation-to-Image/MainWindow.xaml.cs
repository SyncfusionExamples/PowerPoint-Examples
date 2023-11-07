using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Syncfusion.Presentation;
using System.IO;
using System.Drawing;

namespace Convert_PowerPoint_Presentation_to_Image
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
            //Open a PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"../../Data/Input.pptx")))
            {
                //Converts the first slide into image.
                System.Drawing.Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
                //Save the image as jpeg.
                image.Save(Path.GetFullPath(@"../../PPTXtoImage.Jpeg"));
            }
            //Launch the image file.
            if (System.Windows.MessageBox.Show("Do you want to view the image file?", "Image file created",
  MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../PPTXtoImage.Jpeg")) { UseShellExecute = true };
                process.Start();
            }
        }
    }
}

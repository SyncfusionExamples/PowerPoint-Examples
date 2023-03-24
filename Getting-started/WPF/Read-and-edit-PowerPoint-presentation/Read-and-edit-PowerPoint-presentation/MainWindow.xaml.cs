using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
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
using System.Windows.Shapes;

namespace Read_and_edit_PowerPoint_presentation
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
        private void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //Opens an existing PowerPoint presentation.
            IPresentation pptxDoc = Presentation.Open("../../Data/Template.pptx");

            //Gets the first slide from the PowerPoint presentation
            ISlide slide = pptxDoc.Slides[0];

            //Gets the first shape of the slide
            IShape shape = slide.Shapes[0] as IShape;

            //Change the text of the shape
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";

            //Saves the Presentation to the file system.
            pptxDoc.Save("../../Result.pptx");

            //Close the PowerPoint presentation
            pptxDoc.Close();
            MessageBox.Show("PowerPoint Presentation generated successfully");
        }
    }
}

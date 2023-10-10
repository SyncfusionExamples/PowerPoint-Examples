using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;

namespace Convert_PowerPoint_Presentation_to_PDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnConvert_Click(object sender, EventArgs e)
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

                //Launch the PDF file.
                System.Diagnostics.Process.Start(Path.GetFullPath(@"../../Sample.pdf"));
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }


    }
}

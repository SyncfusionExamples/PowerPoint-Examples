using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Read_and_edit_PowerPoint_presentation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnCreate_Click(object sender, EventArgs e)
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

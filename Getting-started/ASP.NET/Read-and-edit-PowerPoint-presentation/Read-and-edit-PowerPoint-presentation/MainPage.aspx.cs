using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Read_and_edit_PowerPoint_presentation
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            //Open an existing PowerPoint presentation
            IPresentation pptxDoc = Presentation.Open(Server.MapPath("~/App_Data/Template.pptx"));

            //Gets the first slide from the PowerPoint presentation
            ISlide slide = pptxDoc.Slides[0];

            //Gets the first shape of the slide
            IShape shape = slide.Shapes[0] as IShape;

            //Change the text of the shape
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";

            //Save the PowerPoint presentation
            pptxDoc.Save("Result.pptx", FormatType.Pptx, Response);

            //Close the PowerPoint presentation
            pptxDoc.Close();
        }
    }
}
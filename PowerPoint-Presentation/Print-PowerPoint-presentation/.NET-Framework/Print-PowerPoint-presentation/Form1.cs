using Syncfusion.Presentation;
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

namespace Print_PowerPoint_presentation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Image[] images = null;
        int startPageIndex;
        int endPageIndex;
        private void Form1_Load(object sender, EventArgs e)
        {
            //Loads or open an PowerPoint Presentation.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Convert the slides to images.
                    images = pptxDoc.RenderAsImages(Syncfusion.Drawing.ImageType.Bitmap);
                    //Initialize the start and page for printing.
                    int startPageIndex = 1;
                    int endPageIndex = images.Length;
                    //Create new PrintDialog instance.
                    using (PrintDialog printDialog = new PrintDialog())
                    {
                        //Set new PrintDocument instance to print dialog.
                        printDialog.Document = new PrintDocument();
                        //Enable the print current page option.
                        printDialog.AllowCurrentPage = true;
                        //Enable the print selected pages option.
                        printDialog.AllowSomePages = true;
                        //Set the start and end page index.
                        printDialog.PrinterSettings.FromPage = 1;
                        printDialog.PrinterSettings.ToPage = images.Length;
                        //Open the print dialog box.
                        if (printDialog.ShowDialog() == DialogResult.OK)
                        {
                            //Check whether the selected page range is valid or not.
                            if (printDialog.PrinterSettings.FromPage > 0 && printDialog.PrinterSettings.ToPage <= images.Length)
                            {
                                //Update the start page of the document to print.
                                startPageIndex = printDialog.PrinterSettings.FromPage - 1;
                                //Update the end page of the document to print.
                                endPageIndex = printDialog.PrinterSettings.ToPage;
                                //Hook the PrintPage event to handle be drawing pages for printing.
                                printDialog.Document.PrintPage += new PrintPageEventHandler(PrintPageMethod);
                                //Print the document.
                                printDialog.Document.Print();
                            }
                        }
                    }
                }
            } 
        }
		/// <summary>
        /// Print the pages.
        /// </summary>
        private void PrintPageMethod(object sender, PrintPageEventArgs e)
        {
            //Get the print start page width.
            int currentPageWidth = images[startPageIndex].Width;
            //Get the print start page height.
            int currentPageHeight = images[startPageIndex].Height;
            //Get the visible bounds width for print.
            int visibleClipBoundsWidth = (int)e.Graphics.VisibleClipBounds.Width;
            //Get the visible bounds height for print.
            int visibleClipBoundsHeight = (int)e.Graphics.VisibleClipBounds.Height;
            //Check whether the page layout is landscape or portrait.
            if (currentPageWidth > currentPageHeight)
            {
                //Translate the position.
                e.Graphics.TranslateTransform(0, visibleClipBoundsHeight);
                //Rotate the object at 270 degrees.
                e.Graphics.RotateTransform(270.0f);
                //Draw the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, currentPageWidth, currentPageHeight));
            }
            else
                //Draw the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, visibleClipBoundsWidth, visibleClipBoundsHeight));
            //Dispose the current page image after drawing.
            images[startPageIndex].Dispose();
            //Increment the start page index.
            startPageIndex++;
            //Update whether the document contains more pages to print or not.
            if (startPageIndex < endPageIndex)
                e.HasMorePages = true;
            else
                startPageIndex = 0;
        }
    }
}
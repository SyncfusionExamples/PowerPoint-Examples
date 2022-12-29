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
            //Loads or open an PowerPoint Presentation
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Converts the slides to images
                    images = pptxDoc.RenderAsImages(Syncfusion.Drawing.ImageType.Bitmap);
                    //Initializes the start and page for printing
                    int startPageIndex = 1;
                    int endPageIndex = images.Length;
                    //Creates new PrintDialog instance.
                    using (PrintDialog printDialog = new PrintDialog())
                    {
                        //Sets new PrintDocument instance to print dialog.
                        printDialog.Document = new PrintDocument();
                        //Enables the print current page option.
                        printDialog.AllowCurrentPage = true;
                        //Enables the print selected pages option.
                        printDialog.AllowSomePages = true;
                        //Sets the start and end page index
                        printDialog.PrinterSettings.FromPage = 1;
                        printDialog.PrinterSettings.ToPage = images.Length;
                        //Opens the print dialog box.
                        if (printDialog.ShowDialog() == DialogResult.OK)
                        {
                            //Checks whether the selected page range is valid or not
                            if (printDialog.PrinterSettings.FromPage > 0 && printDialog.PrinterSettings.ToPage <= images.Length)
                            {
                                //Updates the start page of the document to print.
                                startPageIndex = printDialog.PrinterSettings.FromPage - 1;
                                //Updates the end page of the document to print.
                                endPageIndex = printDialog.PrinterSettings.ToPage;
                                //Hooks the PrintPage event to handle be drawing pages for printing.
                                printDialog.Document.PrintPage += new PrintPageEventHandler(PrintPageMethod);
                                //Prints the document.
                                printDialog.Document.Print();
                            }
                        }
                    }
                }
            } 
        }
        private void PrintPageMethod(object sender, PrintPageEventArgs e)
        {
            //Gets the print start page width.
            int currentPageWidth = images[startPageIndex].Width;
            //Gets the print start page height.
            int currentPageHeight = images[startPageIndex].Height;
            //Gets the visible bounds width for print.
            int visibleClipBoundsWidth = (int)e.Graphics.VisibleClipBounds.Width;
            //Gets the visible bounds height for print.
            int visibleClipBoundsHeight = (int)e.Graphics.VisibleClipBounds.Height;
            //Checks whether the page layout is landscape or portrait.
            if (currentPageWidth > currentPageHeight)
            {
                //Translates the position.
                e.Graphics.TranslateTransform(0, visibleClipBoundsHeight);
                //Rotates the object at 270 degrees
                e.Graphics.RotateTransform(270.0f);
                //Draws the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, currentPageWidth, currentPageHeight));
            }
            else
                //Draws the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, visibleClipBoundsWidth, visibleClipBoundsHeight));
            //Disposes the current page image after drawing.
            images[startPageIndex].Dispose();
            //Increments the start page index.
            startPageIndex++;
            //Updates whether the document contains more pages to print or not.
            if (startPageIndex < endPageIndex)
                e.HasMorePages = true;
            else
                startPageIndex = 0;
        }
    }
}
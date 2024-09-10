using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Presentation;

namespace Based_on_Entrance_animation_effects
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens a PowerPoint Presentation
            IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/AnimationConverter.pptx"));

            //Initialize the PresentationAnimationConverter to perform slide to image conversion based on animation order.
            using (PresentationAnimationConverter animationConverter = new PresentationAnimationConverter())
            {
                int i = 0;
                foreach (ISlide slide in pptxDoc.Slides)
                {
                    //Convert the PowerPoint slide to a series of images based on entrance animation effects.
                    Stream[] imageStreams = animationConverter.Convert(slide, ExportImageFormat.Png);

                    //Save the image stream.
                    foreach (Stream stream in imageStreams)
                    {
                        i++;
                        //Reset the stream position.
                        stream.Position = 0;

                        //Create the output image file stream.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output" + i + ".png")))
                        {
                            //Copy the converted image stream into created output stream.
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}

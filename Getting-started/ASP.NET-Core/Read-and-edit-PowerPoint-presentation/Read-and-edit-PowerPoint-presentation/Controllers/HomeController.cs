﻿using Read_and_edit_PowerPoint_presentation.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Presentation;
using System.Diagnostics;

namespace Read_and_edit_PowerPoint_presentation.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult CreatePowerPoint()
        {
            using FileStream fileStreamPath = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //Open an existing PowerPoint presentation.
            using IPresentation pptxDoc = Presentation.Open(fileStreamPath);
            //Get the first slide from the PowerPoint presentation.
            ISlide slide = pptxDoc.Slides[0];
            //Get the first shape of the slide.
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape.
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Save the PowerPoint Presentation as stream.
            MemoryStream pptxStream = new();
            pptxDoc.Save(pptxStream);
            pptxStream.Position = 0;
            //Download Powerpoint document in the browser.
            return File(pptxStream, "application/powerpoint", "Result.pptx");
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
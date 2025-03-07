using Syncfusion.Presentation;

//Create an instance for PowerPoint
IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);

IShape textboxShape = slide.AddTextBox(200, 50, 150, 50);
IParagraph paragraph = textboxShape.TextBody.AddParagraph();
//Adds a TextPart to the paragraph
ITextPart textPart = paragraph.AddTextPart();
//Adds text to the TextPart
textPart.Text = "Hello World!";
//Access the animation sequence to create effects
ISequence sequence = slide.Timeline.MainSequence;
//Add bounce effect to the shape
IEffect bounceEffect = sequence.AddEffect(textboxShape, EffectType.Bounce, EffectSubtype.None, EffectTriggerType.WithPrevious);


IShape textboxShape1 = slide.AddTextBox(200, 150, 150, 50);
//Adds paragraph to the textbody of textbox
IParagraph paragraph1 = textboxShape1.TextBody.AddParagraph();
//Adds a TextPart to the paragraph
ITextPart textPart1 = paragraph1.AddTextPart();
//Adds text to the TextPart
textPart1.Text = "New Textbox Added";
//Access the animation sequence to create effects
ISequence sequence1 = slide.Timeline.MainSequence;
//Add bounce effect to the shape
IEffect bounceEffect1 = sequence1.AddEffect(textboxShape1, EffectType.Swivel, EffectSubtype.None, EffectTriggerType.WithPrevious);


FileStream pictureStream = new FileStream(@"..\..\..\Data\Image.jpeg", FileMode.Open);
//Adds the picture to a slide by specifying its size and position.
IPicture picture = slide.Pictures.AddPicture(pictureStream, 200, 250, 300, 200);
//Access the animation sequence to create effects
ISequence sequence2 = slide.Timeline.MainSequence;
//Add bounce effect to the shape
IEffect bounceEffect2 = sequence2.AddEffect(picture as IShape, EffectType.Bounce, EffectSubtype.None, EffectTriggerType.WithPrevious);


//Save the PowerPoint Presentation as stream
FileStream outputStream = new FileStream(@"..\..\..\Data\Sample.pptx", FileMode.Create);
pptxDoc.Save(outputStream);

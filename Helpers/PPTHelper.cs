using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;
namespace Auxi_PowerPointEdit.Helpers
{
    public static class Constants
    {
        public static string _INPUTSLIDE_ = "Input Slide";
        public static string _OUTPUTSLIDE_ = "Output Slide";
        public static string _TEXTBOX_ = "TextBox";
        public static string _ARROW_ = "Arrow";
    }
    public class PPTHelper
    {
        private static PresentationDocument Presentation;
        public static int SlideRef = 0;

        public void OpenPptxFile(string FilePath)
        {
            Presentation = PresentationDocument.Open(FilePath, true);
        }
        public void Dispose()
        {
            Presentation.Dispose();
        }

        public SlidePart CloneInputSlide()
        {
            if (Presentation == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            PresentationPart presentationPart = Presentation.PresentationPart;

            SlideId slideId = presentationPart.Presentation.SlideIdList.GetFirstChild<SlideId>();

            string relId = slideId.RelationshipId;

            // Get the slide part by the relationship ID.

            SlidePart inputSlide = (SlidePart)presentationPart.GetPartById(relId);

            if (inputSlide == default(SlidePart))
            {
                throw new ArgumentException("SlidePart");
            }
            //Create a new slide part in the presentation. 
            SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>("OutPutSlideResult-"+SlideRef);
            SlideRef++;
            //Add the slide template content into the new slide. 

            newSlidePart.FeedData(inputSlide.GetStream(FileMode.Open));
            //Make sure the new slide references the proper slide layout. 
            newSlidePart.AddPart(inputSlide.SlideLayoutPart);
            //Get the list of slide ids. 
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            //Deternmine where to add the next slide (find max number of slides). 
            uint maxSlideId = 1;
            SlideId prevSlideId = null;
            foreach (SlideId slideID in slideIdList.ChildElements)
            {
                if (slideID.Id > maxSlideId)
                {
                    maxSlideId = slideID.Id;
                    prevSlideId = slideID;
                }
            }
            maxSlideId++;
            //Add the new slide at the end of the deck. 
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            //Make sure the id and relid are set appropriately. 
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
            return newSlidePart;
        }

        public static void FormatTitle(SlidePart slidePart)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slidePart.Slide != null)
            {

                Shape titleShape = slidePart.Slide.Descendants<Shape>().Where(d => IsTitleShape(d) && d.InnerText.Contains(Constants._INPUTSLIDE_)).FirstOrDefault();
                if (titleShape != default(Shape))
                {

                    D.Paragraph paragraph = titleShape.TextBody.Descendants<D.Paragraph>().FirstOrDefault();
                    if (paragraph != default(D.Paragraph))
                    {
                        D.Run run = paragraph.Descendants<D.Run>().FirstOrDefault();
                        run.Text = new D.Text(Constants._OUTPUTSLIDE_);
                        paragraph.InsertAt(new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center }, 0);
                        run.RunProperties.InsertAt(new D.ComplexScriptFont() { CharacterSet = -78, PitchFamily = 2, Typeface = "Beirut" }, 0);
                        run.RunProperties.InsertAt(new D.LatinFont() { CharacterSet = -78, PitchFamily = 2, Typeface = "Beirut" }, 0);
                    }
                }
            }
        }

        public static void FormatFLow(SlidePart slidePart)
        {
            if (slidePart != default(SlidePart))
            {
                IEnumerable<Shape> flowChartShapes = slidePart.Slide.Descendants<Shape>().Where(d => d.Descendants<NonVisualDrawingProperties>().Where(e => e.Name.InnerText.Contains(Constants._ARROW_)).Count() > 0);
                int i = 0;
                long firstId = 0;
                D.Transform2D prevBox = new D.Transform2D();
                foreach (Shape shape in flowChartShapes)
                {
                    String title = String.Empty;
                    NonVisualDrawingProperties meta = shape.Descendants<NonVisualDrawingProperties>().FirstOrDefault();


                    //resize shape
                    D.Transform2D currentShapePos = shape.Descendants<D.Transform2D>().FirstOrDefault();

                    currentShapePos.Extents.Cx = 2775204;
                    currentShapePos.Extents.Cy = 1446936;
                    if (i == 0)
                    {
                        firstId = meta.Id;
                        currentShapePos.Offset.X = 458569;
                        currentShapePos.Offset.Y = 1712339;

                    }
                    else
                    {
                        currentShapePos.Offset.X = Convert.ToInt64((prevBox.Extents.Cx.Value - (prevBox.Extents.Cx.Value * 0.20)) + prevBox.Offset.X);
                        currentShapePos.Offset.Y = prevBox.Offset.Y;
                    }



                    prevBox = currentShapePos;


                    if (meta != default(NonVisualDrawingProperties))
                    {
                        long idToFind = (firstId + i) - (flowChartShapes.Count() + 1);//check 

                        Shape titleShape = slidePart.Slide.Descendants<Shape>().Where(d => d.Descendants<NonVisualDrawingProperties>().First() != default(NonVisualDrawingProperties) &&
                                                                                          d.Descendants<NonVisualDrawingProperties>().First().Id == idToFind &&
                                                                                          d.Descendants<NonVisualDrawingProperties>().First().Name.InnerText.Contains(Constants._TEXTBOX_))
                                                                                          .FirstOrDefault();
                        if (titleShape != default(Shape))
                        {
                            D.Paragraph data = titleShape.Descendants<D.Paragraph>().FirstOrDefault();

                            title = data != default(D.Paragraph) ? data.InnerText : title;
                        }

                        //slidePart.Slide.RemoveChild(titleShape);
                        titleShape.Remove();
                        slidePart.Slide.Save();

                        D.Paragraph paragraph = shape.Descendants<D.Paragraph>().FirstOrDefault();
                        if (paragraph != default(D.Paragraph))
                        {
                            paragraph.RemoveAllChildren();
                            slidePart.Slide.Save();
                            D.Run run1 = new D.Run();
                            paragraph.AppendChild(new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center });
                            run1.AppendChild(new D.RunProperties() { Language = "en-US", Dirty = true });
                            run1.Text = new D.Text(title);
                            paragraph.AppendChild(run1);
                        }
                        slidePart.Slide.Save();
                    }
                    i++;

                }
            }

        }

        public static void FormatBlPoint(SlidePart slidePart)
        {
            IEnumerable<Shape> bulletPointShapes = slidePart.Slide.Descendants<Shape>().Where(d => d.Descendants<NonVisualDrawingProperties>().Where(e => e.Name.InnerText.Contains(Constants._TEXTBOX_)).Count() > 0);
            foreach (Shape shape in bulletPointShapes)
            {
                IEnumerable<D.Paragraph> paragraphs = shape.Descendants<D.Paragraph>().Where(d => d.Descendants<D.ParagraphProperties>().Count() > 0);
                foreach (D.Paragraph paragraph in paragraphs)
                {
                    D.ParagraphProperties paragProps = paragraph.Descendants<D.ParagraphProperties>().FirstOrDefault();
                    if (paragProps != default(D.ParagraphProperties))
                    {

                        D.BulletFont bulletFont = paragProps.Descendants<D.BulletFont>().FirstOrDefault();
                        if (bulletFont != default(D.BulletFont))
                        {

                            bulletFont.CharacterSet = 0;
                            bulletFont.PitchFamily = 34;
                            bulletFont.Typeface = "Arial";
                        }

                        D.CharacterBullet characterBullet = paragProps.Descendants<D.CharacterBullet>().FirstOrDefault();

                        if (characterBullet != default(D.CharacterBullet))
                        {
                            characterBullet.Char = new DocumentFormat.OpenXml.StringValue("•");

                        }
                        else
                        {
                            D.AutoNumberedBullet autoNumberedBullet= paragProps.Descendants<D.AutoNumberedBullet>().FirstOrDefault();
                            if(autoNumberedBullet!=default(D.AutoNumberedBullet))
                            {
                                autoNumberedBullet.Remove();
                                slidePart.Slide.Save();
                            }
                            characterBullet = new D.CharacterBullet() { Char = new DocumentFormat.OpenXml.StringValue("•") };
                            paragProps.AppendChild(characterBullet);
                        }
                        D.Run run = paragraph.Descendants<D.Run>().FirstOrDefault();
                        run.RunProperties.Italic = false;
                        run.RunProperties.Bold = false;
                        run.RunProperties.Underline = D.TextUnderlineValues.None;
                    }

                }
            }
            slidePart.Slide.Save();
        }

        private static bool IsTitleShape(Shape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape.
                    case PlaceholderValues.Title:

                    // A centered title.
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }
            return false;
        }

    }
}

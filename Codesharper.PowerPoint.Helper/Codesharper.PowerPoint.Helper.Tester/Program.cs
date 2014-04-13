﻿namespace Codesharper.PowerPoint.Helper.Tester
{
    using System.Diagnostics;

    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Implementations;
    using Codesharper.PowerPoint.Helper.Objects;

    using Microsoft.Office.Core;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    internal class Program
    {
        private const string presentationFile = @"c:\temp\testPPT.ppt";

        /// <summary>
        ///     Utility method to determine if a number if odd or even
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsOdd(int value)
        {
            return value % 2 != 0;
        }

        private static void Main(string[] args)
        {
            // Setup Objects
            var pptApplicationManager = new PowerPointApplicationManager();
            var pptSlideManager = new SlideManager();
            var pptShapeManager = new ShapesManager();
            var pptPresentationManager = new PresentationManager();

            // Step 1. Create a PowerPoint application instance
            PPT.Application pptApplication = pptApplicationManager.CreatePowerPointApplication();

            // Step 2. Create a new PowerPoint presentation
            PPT.Presentation pptPresentation = pptPresentationManager.CreatePowerPointPresentation(
                    pptApplication,
                    false);

            // Step 3. Add a slide to the end of the presentation
            PPT.Slide lastSlide = pptSlideManager.AddSlideToEnd(pptPresentation);

            // Step 4. Add a slide to the start of the presentation
            PPT.Slide firstSlide = pptSlideManager.AddSlideToStart(pptPresentation);

            // Step 5. Add a slide at position 1 in the presentation
            PPT.Slide pos1Slide = pptSlideManager.InsertSlide(pptPresentation, 1);

            // Step 6a. Add a text box to each of the slides showing the slide number
            foreach (PPT.Slide slide in pptPresentation.Slides)
            {
                PPT.Shape textboxShape = pptShapeManager.AddTextBoxToSlide(
                        slide,
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        200f,
                        200f,
                        200f,
                        200f);

                pptShapeManager.SetTextBoxText(textboxShape, "This is slide " + slide.SlideIndex);

                // Step 6b. Add a shape to the slide; star if its an odd numbered slide, rectanle if its an even numbered slide
                if (IsOdd(slide.SlideIndex))
                {
                    PPT.Shape starShape = pptShapeManager.DrawShape(
                            slide,
                            MsoAutoShapeType.msoShape8pointStar,
                            400f,
                            150f,
                            250f,
                            250f);
                }
                else
                {
                    PPT.Shape rectangleSHape = pptShapeManager.DrawShape(
                            slide,
                            MsoAutoShapeType.msoShapeRectangle,
                            400f,
                            150f,
                            250f,
                            250f);
                }

                // Step 6c. Add a comment to the slide telling you if the index of the slide is odd or even
                string slideCommentText = string.Empty;

                if (IsOdd(slide.SlideIndex))
                {
                    slideCommentText = @"This is an odd numbered slide";
                }
                else
                {
                    slideCommentText = @"This is an even numbered slide";
                }

                PPT.Slide slideComment = pptSlideManager.AddComment(
                        slide,
                        new SlideComment
                            {
                                    Author = "Test Author",
                                    AuthorInitials = "TA",
                                    Comment = slideCommentText,
                                    LeftPosition = 200f,
                                    TopPosition = 50f
                            });
            }

            // Step 7. Clone the first slide and make it the last slide
            PPT.SlideRange cloneSlide = pptSlideManager.CloneSlide(pptPresentation, firstSlide, Locations.Location.Last);

            // Step 8a. Add a new 'summary' slide to the end of the presentation
            PPT.Slide summarySlide = pptSlideManager.AddSlideToEnd(pptPresentation);

            // Step 8b. Add a count to the slide showing the total number of slides in the presentation
            PPT.Shape summarySlideCountText = pptShapeManager.AddTextBoxToSlide(
                    summarySlide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    300f,
                    300f,
                    300f,
                    300f);

            // Step 8c. Set the text to the number of slides in the presentation
            pptShapeManager.SetTextBoxText(
                    summarySlideCountText,
                    "There are " + pptSlideManager.GetSlideCount(pptPresentation).ToString() + " slides in this presentation!");

            // Move the summary slide to be the first slide in the presentation
            pptSlideManager.MoveSlide(pptPresentation, summarySlide, Locations.Location.First);

            // Step 9. Save the presentation to c:\temp\testPPT.pptx and open it
            pptPresentationManager.SavePresentationAs(
                    pptPresentation,
                    presentationFile,
                    PPT.PpSaveAsFileType.ppSaveAsPresentation,
                    true);
            pptPresentationManager.ClosePresentation(pptPresentation);

            Process.Start(presentationFile);
        }
    }
}
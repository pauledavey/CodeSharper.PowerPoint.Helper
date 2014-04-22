namespace Codesharper.PowerPoint.Helper.Tester
{
    #region Using Directives

    using System.Collections.Generic;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Implementations;
    using Codesharper.PowerPoint.Helper.Objects;
    using Codesharper.PowerPoint.Helper.Tester.Demonstrations;

    using Microsoft.Office.Core;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    internal class Program
    {
        private const string presentationFile = @"c:\temp\testPPT.pptx";

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
            Demo1.CreateDemonstrationPresentation();



            //// Setup Objects
            //var pptApplicationManager = new PowerPointApplicationManager();
            //var pptSlideManager = new SlideManager();
            //var pptShapeManager = new ShapesManager();
            //var pptPresentationManager = new PresentationManager();
            //var pptChartManager = new ChartManager();

            //// Step 1. Create a PowerPoint application instance
            //PPT.Application pptApplication = pptApplicationManager.CreatePowerPointApplication();

            //// Step 2. Create a new PowerPoint presentation
            //PPT.Presentation pptPresentation = pptPresentationManager.CreatePowerPointPresentation(pptApplication, true);

            //// Step 3. Add a slide to the end of the presentation
            //PPT.Slide lastSlide = pptSlideManager.AddSlideToEnd(pptPresentation);

            //// Step 4. Add a slide to the start of the presentation
            //PPT.Slide firstSlide = pptSlideManager.AddSlideToStart(pptPresentation);

            //// Step 5. Add a slide at position 1 in the presentation
            //PPT.Slide pos1Slide = pptSlideManager.InsertSlide(pptPresentation, 1);

            //// Step 6a. Add a text box to each of the slides showing the slide number
            //foreach (PPT.Slide slide in pptPresentation.Slides)
            //{
            //    PPT.Shape textboxShape = pptShapeManager.AddTextBoxToSlide(
            //            slide,
            //            MsoTextOrientation.msoTextOrientationHorizontal,
            //            200f,
            //            200f,
            //            200f,
            //            200f);

            //    pptShapeManager.SetTextBoxText(textboxShape, "This is slide " + slide.SlideIndex);

            //    // Step 6b. Add a shape to the slide; star if its an odd numbered slide, rectanle if its an even numbered slide
            //    if (IsOdd(slide.SlideIndex))
            //    {
            //        PPT.Shape starShape = pptShapeManager.DrawShape(
            //                slide,
            //                MsoAutoShapeType.msoShape8pointStar,
            //                400f,
            //                150f,
            //                250f,
            //                250f);
            //    }
            //    else
            //    {
            //        PPT.Shape rectangleSHape = pptShapeManager.DrawShape(
            //                slide,
            //                MsoAutoShapeType.msoShapeRectangle,
            //                400f,
            //                150f,
            //                250f,
            //                250f);
            //    }

            //    // Step 6c. Add a comment to the slide telling you if the index of the slide is odd or even
            //    string slideCommentText = string.Empty;

            //    if (IsOdd(slide.SlideIndex))
            //    {
            //        slideCommentText = @"This is an odd numbered slide";
            //    }
            //    else
            //    {
            //        slideCommentText = @"This is an even numbered slide";
            //    }

            //    PPT.Slide slideComment = pptSlideManager.AddComment(
            //            slide,
            //            new SlideComment
            //                {
            //                        Author = "Test Author",
            //                        AuthorInitials = "TA",
            //                        Comment = slideCommentText,
            //                        LeftPosition = 200f,
            //                        TopPosition = 50f
            //                });
            }

        //    // Step 7. Clone the first slide and make it the last slide
        //    PPT.SlideRange cloneSlide = pptSlideManager.CloneSlide(pptPresentation, firstSlide, Locations.Location.Last);

        //    // Step 8a. Add a new 'summary' slide to the end of the presentation
        //    PPT.Slide summarySlide = pptSlideManager.AddSlideToEnd(pptPresentation);

        //    // Step 8b. Add a count to the slide showing the total number of slides in the presentation
        //    PPT.Shape summarySlideCountText = pptShapeManager.AddTextBoxToSlide(
        //            summarySlide,
        //            MsoTextOrientation.msoTextOrientationHorizontal,
        //            300f,
        //            300f,
        //            300f,
        //            300f);

        //    // Step 8c. Set the text to the number of slides in the presentation
        //    pptShapeManager.SetTextBoxText(
        //            summarySlideCountText,
        //            "There are " + pptSlideManager.GetSlideCount(pptPresentation).ToString()
        //            + " slides in this presentation!");

        //    // Step9. Move the summary slide to be the first slide in the presentation
        //    pptSlideManager.MoveSlide(pptPresentation, summarySlide, Locations.Location.First);

        //    // Step10. Set transition effects for each slide in the presentation
        //    foreach (PPT.Slide currSlide in pptPresentation.Slides)
        //    {
        //        pptSlideManager.SetSlideTransition(
        //                currSlide,
        //                PPT.PpEntryEffect.ppEffectBlindsVertical,
        //                PPT.PpTransitionSpeed.ppTransitionSpeedMedium);
        //    }

        //    // Step11. Lets do some Charting!!
        //    lastSlide = pptSlideManager.AddSlideToEnd(pptPresentation);

        //    // Defining Chart Data
        //    var columns = new string[] { "Bananas", "Apples", "Oranges", "Pears", "Grapes" };
        //    var chartSeries1 = new ChartSeries()
        //                           {
        //                                   name = "Pauls Series",
        //                                   seriesData = new string[] { "100", "200", "300", "400", "50" },
        //                                   seriesType = XlChartType.xlLine
        //                           };

        //    var chartSeries2 = new ChartSeries()
        //                           {
        //                                   name = "Codesharper Series",
        //                                   seriesData = new string[] { "1000", "2000", "3000", "3050", "2200" },
        //                                   seriesType = XlChartType.xlArea
        //                           };

        //    var seriesData = new List<ChartSeries> { chartSeries1, chartSeries2 };

        //    var chart = pptChartManager.CreateChart(lastSlide, columns, seriesData);

        //    var chartSeries3 = new ChartSeries()
        //    {
        //        name = "Codesharper2 Series",
        //        seriesData = new string[] { "2500", "3500", "4500", "5500", "6500" },
        //        seriesType = XlChartType.xlLine
        //    };

        //    pptChartManager.AddSeriesToExistingChart(chart, chartSeries3);

        //    // var series = pptChartManager.GetChartSeriesByName(chart, "Codesharper2 Series");

           


        //    pptChartManager.AddChartTitle(
        //            chart,
        //            new ChartTitle()
        //                {
        //                        bold = true,
        //                        fontSize = 14,
        //                        italic = false,
        //                        titleText = "Welcome to my Chart!",
        //                        underline = false
        //                });

        //    pptChartManager.AddChartLegend(
        //            chart,
        //            new ChartLegend() { bold = true, fontSize = 14, italic = false, underline = false });

        //    // Step 99. Save the presentation to c:\temp\testPPT.pptx and open it
        //    pptPresentationManager.SavePresentationAs(
        //            pptPresentation,
        //            presentationFile,
        //            PPT.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
        //            true);

        //    // Step 100a. Export the first slide in the presentation
        //    pptSlideManager.Export(pptPresentation.Slides[1], @"C:\temp\firstslide.png", ImageFormats.Formats.png);

        //    // Step 100b. Export all slides in the presentation as PNG
        //    pptSlideManager.ExportAll(pptPresentation, @"C:\temp\", ImageFormats.Formats.png);

        //    // Step 100c. Export all slides in the presentation as PNG
        //    pptSlideManager.ExportAll(pptPresentation, @"C:\temp\", ImageFormats.Formats.jpg);

        //    // Step 100d. Export all slides in the presentation as PNG
        //    pptSlideManager.ExportAll(pptPresentation, @"C:\temp\", ImageFormats.Formats.bmp);

        //    pptPresentationManager.ClosePresentation(pptPresentation);
        //    pptApplicationManager.ClosePowerPointApplication(pptApplication);
        //}
    }
}
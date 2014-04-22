using Codesharper.PowerPoint.Helper.Tester;

namespace Codesharper.PowerPoint.Helper.Tester
{
    #region Using Directives

    using System;
    using System.Collections.Generic;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    internal class Program
    {
        private const string ResourceFolder = @"\CodeSharper Template\";

        private const string Slide1Graphic = @"Slide1Graphic.png";

        private const string TemplatePath = @"template.pptx";

        private const string OutputFile = @"Demonstration1.pptx";

        // Setup Objects
        private static readonly PowerPointApplicationManager PptApplicationManager = new PowerPointApplicationManager();

        private static readonly SlideManager PptSlideManager = new SlideManager();

        private static readonly ShapesManager PptShapeManager = new ShapesManager();

        private static readonly PresentationManager PptPresentationManager = new PresentationManager();

        private static readonly ChartManager PptChartManager = new ChartManager();

        private static void Main(string[] args)
        {

            // Step1. Create a PowerPoint application instance
            var powerPointApplication = PptApplicationManager.CreatePowerPointApplication();

            // Step2. Open an existing PowerPoint presentation
            var pptPresentation = PptPresentationManager.OpenExistingPowerPointPresentation(
                    powerPointApplication,
                    Environment.CurrentDirectory + ResourceFolder + TemplatePath);

            // Step3. Create two clones of our second slide (we will use this a lot)
            PptSlideManager.CloneSlide(pptPresentation, pptPresentation.Slides[2], Locations.Location.Last);
            PptSlideManager.CloneSlide(pptPresentation, pptPresentation.Slides[2], Locations.Location.Last);

            // Step4. Decorate Slides 1,2,3,4
            DecorateSlides(pptPresentation);

            // Step 5. Configure Transitions for slides in the presentation
            PptSlideManager.SetSlideTransition(
                    pptPresentation.Slides[1],
                    PPT.PpEntryEffect.ppEffectWheel8Spokes,
                    PPT.PpTransitionSpeed.ppTransitionSpeedMedium);
            PptSlideManager.SetSlideTransition(
                    pptPresentation.Slides[2],
                    PPT.PpEntryEffect.ppEffectBlindsHorizontal,
                    PPT.PpTransitionSpeed.ppTransitionSpeedMedium);
            PptSlideManager.SetSlideTransition(
                    pptPresentation.Slides[3],
                    PPT.PpEntryEffect.ppEffectShredRectangleOut,
                    PPT.PpTransitionSpeed.ppTransitionSpeedMedium);
            PptSlideManager.SetSlideTransition(
                    pptPresentation.Slides[4],
                    PPT.PpEntryEffect.ppEffectGlitterDiamondUp,
                    PPT.PpTransitionSpeed.ppTransitionSpeedMedium);

            // Step 6. Run Cleanup (Save presentation, export slides as PNG and dispose of objects
            Cleanup(pptPresentation, powerPointApplication);

            // Step 7. Run the presentation
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + ResourceFolder);
        }

        private static void SetSlideFooter(PPT.Slide slide, int slideNumber)
        {
            // Insert the textboxes
            var textShape1 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    905f,
                    512f,
                    100f,
                    40f);

            // Set the text in the textboxes
            PptShapeManager.SetTextBoxText(textShape1, "slide " + slideNumber);
            textShape1.TextEffect.FontBold = MsoTriState.msoFalse;
            textShape1.TextEffect.FontSize = 14f;
        }

        private static void DecorateSlides(PPT.Presentation pptPresentation)
        {
            DecorateSlideOne(pptPresentation.Slides[1]);
            DecorateSlideTwo(pptPresentation.Slides[2]);
            DecorateSlideThree(pptPresentation.Slides[3]);
            DecorateSlideFour(pptPresentation.Slides[4]);
        }

        private static void DecorateSlideOne(PPT.Slide slide)
        {
            // Insert the graphic for the slide
            var pictureShape = PptShapeManager.AddPicture(
                    slide,
                    Environment.CurrentDirectory + ResourceFolder + Slide1Graphic,
                    100f,
                    220f,
                    200f,
                    200f);

            // Insert the textboxes
            var textShape1 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    320f,
                    280f,
                    400f,
                    40f);

            var textShape2 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    320f,
                    320f,
                    400f,
                    40f);

            // Set the text in the textboxes
            PptShapeManager.SetTextBoxText(textShape1, "CodeSharper PowerPoint Helper");
            textShape1.TextEffect.FontBold = MsoTriState.msoTrue;
            textShape1.TextEffect.FontSize = 42f;

            PptShapeManager.SetTextBoxText(textShape2, "A Demonstration");
            textShape2.TextEffect.FontBold = MsoTriState.msoTrue;
            textShape1.TextEffect.FontSize = 28f;

            SetSlideFooter(slide, 1);
        }

        private static void DecorateSlideTwo(PPT.Slide slide)
        {
            // Insert the textboxes
            var textShape1 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    100f,
                    100f,
                    800f,
                    40f);

            // Set the text in the textboxes
            PptShapeManager.SetTextBoxText(
                    textShape1,
                    "This PowerPoint presentation helps to show you what you can achieve with the CodeSharper PowerPoint helper library."
                    + "\nWe hope that you will find this library helpful."
                    + "\nFeel free to contact email any questions to" + "\nPaul.Davey@Codesharper.co.uk");
            textShape1.TextEffect.FontSize = 16f;
            SetSlideFooter(slide, 2);
        }

        private static void DecorateSlideThree(PPT.Slide slide)
        {
            var columnsList = new string[] { "C#", "VB.Net", "Perl", "Python", "Java" };

            var series1 = new ChartSeries()
                              {
                                      name = "Blog Statistics",
                                      seriesData = new string[] { "8200", "3900", "890", "300", "3278" },
                                      seriesType = XlChartType.xl3DColumn
                              };

            var chartData = new List<ChartSeries>() { series1 };
            var chart = PptChartManager.CreateChart(XlChartType.xlColumnStacked, slide, columnsList, chartData);

            var chartTitle = new ChartTitle()
                                 {
                                         bold = true,
                                         italic = false,
                                         fontSize = 40,
                                         titleText = "Users by Software Language",
                                         underline = false
                                 };

            PptChartManager.AddChartTitle(chart, chartTitle);
            SetSlideFooter(slide, 3);
        }

        private static void DecorateSlideFour(PPT.Slide slide)
        {
            var starShape = PptShapeManager.DrawShape(slide, MsoAutoShapeType.msoShape16pointStar, 45f, 25f, 150f, 150f);
            var upArrowShape = PptShapeManager.DrawShape(
                    slide,
                    MsoAutoShapeType.msoShapeUpArrow,
                    250f,
                    100f,
                    150f,
                    250f);
            var downArrowShape = PptShapeManager.DrawShape(
                    slide,
                    MsoAutoShapeType.msoShapeDownArrow,
                    500f,
                    100f,
                    150f,
                    250f);

            starShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gold);
            starShape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            // Insert the textboxes
            var textShape0 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    70f,
                    75f,
                    100f,
                    100f);
            var textShape1 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    100f,
                    195f,
                    200f,
                    300f);
            var textShape2 = PptShapeManager.AddTextBoxToSlide(
                    slide,
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    650f,
                    120f,
                    200f,
                    300f);

            // Set the text in the textboxes
            PptShapeManager.SetTextBoxText(textShape0, "Important" + "\nFact!");
            textShape0.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            textShape0.Rotation = -45f;
            textShape0.TextFrame.TextRange.Font.Color.RGB =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            textShape0.TextEffect.FontSize = 20f;

            PptShapeManager.SetTextBoxText(
                    textShape1,
                    "The number of C# developers has grown" + "\nby 25% over the past quarter.");
            textShape1.TextEffect.FontSize = 18f;

            PptShapeManager.SetTextBoxText(
                    textShape2,
                    "The number of VB.Net developers has declined" + "\nby 40% over the past quarter.");
            textShape2.TextEffect.FontSize = 18f;

            SetSlideFooter(slide, 4);
        }

        private static void Cleanup(PPT.Presentation presentation, PPT.Application powerPointApplication)
        {
            PptPresentationManager.SavePresentationAs(
                    presentation,
                    Environment.CurrentDirectory + ResourceFolder + OutputFile,
                    PPT.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    true);

            // Step 100b. Export all slides in the presentation as PNG
            PptSlideManager.ExportAll(
                    presentation,
                    Environment.CurrentDirectory + ResourceFolder,
                    ImageFormats.Formats.png);


            PptPresentationManager.ClosePresentation(presentation);
            PptApplicationManager.ClosePowerPointApplication(powerPointApplication);
        }
    }
}

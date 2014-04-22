namespace Codesharper.PowerPoint.Helper.Tester.Demonstrations
{
    #region Using Directives

    using System;

    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    #endregion

    public static class Demo1
    {
        // Setup Objects
        private const string ResourceFolder = @"\CodeSharper Template\";

        private const string Slide1Graphic = @"Slide1Graphic.png";

        private const string TemplatePath = @"template.pptx";

        private static readonly PowerPointApplicationManager PptApplicationManager = new PowerPointApplicationManager();

        private static readonly PresentationManager PptPresentationManager = new PresentationManager();

        private static readonly ShapesManager PptShapeManager = new ShapesManager();

        private static readonly SlideManager PptSlideManager = new SlideManager();

        public static void CreateDemonstrationPresentation()
        {
            // Use a resource PowerPoint presentation as a template
            var templateFile = Environment.CurrentDirectory + ResourceFolder + TemplatePath;

            // Create PowerPoint application instance
            var pptApplication = PptApplicationManager.CreatePowerPointApplication();

            // Open template PowerPoint presentation file
            var pptPresentation = PptPresentationManager.OpenExistingPowerPointPresentation(
                    pptApplication,
                    templateFile);

            // Configure each of our slides as we need them.
            DecorateSlide_One(pptPresentation.Slides[1]);

            PptPresentationManager.SavePresentationAs(
                    pptPresentation,
                    Environment.CurrentDirectory + ResourceFolder + @"Demonstration.pptx",
                    PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    true);
            PptPresentationManager.ClosePresentation(pptPresentation);
            PptApplicationManager.ClosePowerPointApplication(pptApplication);
        }

        private static void DecorateSlide_One(Slide slide)
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

        private static void SetSlideFooter(Slide slide, int slideNumber)
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
    }
}
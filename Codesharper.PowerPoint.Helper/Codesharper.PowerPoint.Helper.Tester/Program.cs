using PPTHelper = Codesharper.PowerPoint.Helper.Implementations.Presentation;
using PPTApplication = Codesharper.PowerPoint.Helper.Implementations.PowerPointApplication;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace Codesharper.PowerPoint.Helper.Tester
{
    using System;

    class Program
    {
        static void Main(string[] args)
        {
            var helper = new PPTHelper();
            var pptAppHelper = new PPTApplication();
            var pptApp = pptAppHelper.CreatePowerPointApplication();

            var presentation = helper.CreatePowerPointPresentation(pptApp);
            var slideAtEnd = helper.AddSlideAtEndOfPresentation(presentation);
            helper.InsertSlideIntoPresentation(presentation, (helper.GetSlideCountInPresentation(presentation) + 1));
            helper.SavePresentationAs(presentation, @"c:\example.ppt", PPT.PpSaveAsFileType.ppSaveAsPresentation, true);

            helper.ClosePresentation(presentation);

            presentation = helper.OpenExistingPowerPointPresentation(pptApp, @"c:\example.ppt");
            slideAtEnd = helper.AddSlideAtEndOfPresentation(presentation);
            helper.InsertSlideIntoPresentation(presentation, (helper.GetSlideCountInPresentation(presentation) + 1));
            helper.SavePresentationAs(presentation, @"c:\example.ppt", PPT.PpSaveAsFileType.ppSaveAsPresentation, true);

            helper.ClosePresentation(presentation);
        }
    }
}

namespace Codesharper.PowerPoint.Helper.Contracts
{
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    public interface IPresentation
    {

        PPT.Application CreatePowerPointApplication();

        PPT.Presentation CreatePowerPointPresentation(PPT.Application powerPointApplication);

        PPT.Presentation OpenExistingPowerPointPresentation(PPT.Application powerPointApplication, string pathAndFileName);

        PPT.Slide AddSlideAtEndOfPresentation(PPT.Presentation presentationToAddSlideTo);

        PPT.Slide InsertSlideIntoPresentation(PPT.Presentation presentationToAddSlideTo, int indexOfSlide);

        int GetSlideCountInPresentation(PPT.Presentation presentation);

        void SavePresentationAs(PPT.Presentation presentationToSave, string pathAndFileName, PPT.PpSaveAsFileType fileType , bool embedTrueTypeFonts = true);

        void ClosePresentation(PPT.Presentation presentationToClose);
    }
}

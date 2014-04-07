namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    #endregion

    public interface IPresentation
    {
        PPT.Slide AddSlideAtEndOfPresentation(PPT.Presentation presentationToAddSlideTo);

        void ClosePresentation(PPT.Presentation presentationToClose);

        PPT.Presentation CreatePowerPointPresentation(PPT.Application powerPointApplication, bool showWindow);

        int GetSlideCountInPresentation(PPT.Presentation presentation);

        PPT.Slide InsertSlideIntoPresentation(PPT.Presentation presentationToAddSlideTo, int indexOfSlide);

        PPT.Presentation OpenExistingPowerPointPresentation(
                PPT.Application powerPointApplication,
                string pathAndFileName);

        void SavePresentationAs(
                PPT.Presentation presentationToSave,
                string pathAndFileName,
                PPT.PpSaveAsFileType fileType,
                bool embedTrueTypeFonts = true);
    }
}
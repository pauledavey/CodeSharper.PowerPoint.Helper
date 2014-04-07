namespace Codesharper.PowerPoint.Helper.Implementations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Contracts;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    #endregion

    public class Presentation : IPresentation
    {
        private const OFFICE.MsoTriState oFalse = OFFICE.MsoTriState.msoFalse;

        private const OFFICE.MsoTriState oTrue = OFFICE.MsoTriState.msoTrue;

        public PPT.Slide AddSlideAtEndOfPresentation(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(
                    (presentationToAddSlideTo.Slides.Count + 1),
                    PPT.PpSlideLayout.ppLayoutBlank);
        }

        public void ClosePresentation(PPT.Presentation presentationToClose)
        {
            presentationToClose.Close();
        }

        public PPT.Presentation CreatePowerPointPresentation(PPT.Application powerPointApplication)
        {
            return powerPointApplication.Presentations.Add(oFalse);
        }

        public int GetSlideCountInPresentation(PPT.Presentation presentation)
        {
            return presentation.Slides.Count;
        }

        public PPT.Slide InsertSlideIntoPresentation(PPT.Presentation presentationToAddSlideTo, int indexOfSlide)
        {
            return presentationToAddSlideTo.Slides.Add(indexOfSlide, PPT.PpSlideLayout.ppLayoutBlank);
        }

        public PPT.Presentation OpenExistingPowerPointPresentation(
                PPT.Application powerPointApplication,
                string pathAndFileName)
        {
            return powerPointApplication.Presentations.Open(pathAndFileName, oFalse, oFalse, oFalse);
        }

        public void SavePresentationAs(
                PPT.Presentation presentationToSave,
                string pathAndFileName,
                PPT.PpSaveAsFileType fileType,
                bool embedTrueTypeFonts)
        {
            if (embedTrueTypeFonts)
            {
                presentationToSave.SaveAs(pathAndFileName, fileType, OFFICE.MsoTriState.msoTrue);
                return;
            }

            presentationToSave.SaveAs(pathAndFileName, fileType, OFFICE.MsoTriState.msoFalse);
        }

        public PPT.Slide FindSlideByItsID(PPT.Presentation presentation, int slideId)
        {
            return presentation.Slides.FindBySlideID(slideId);
        }

        
    }
}
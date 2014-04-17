namespace Codesharper.PowerPoint.Helper.Implementations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Contracts;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    #endregion

    public class PresentationManager : IPresentationManager
    {
        private const OFFICE.MsoTriState oFalse = OFFICE.MsoTriState.msoFalse;

        private const OFFICE.MsoTriState oTrue = OFFICE.MsoTriState.msoTrue;

        /// <summary>
        /// A PowerPoint presentation to close
        /// </summary>
        /// <param name="presentationToClose">Presentation to close</param>
        public void ClosePresentation(PPT.Presentation presentationToClose)
        {
            presentationToClose.Close();
        }

        /// <summary>
        /// Create a PowerPoint presentation
        /// </summary>
        /// <param name="powerPointApplication">An instance of a PPT.Application object</param>
        /// <param name="showPowerPoint">Show the Presentation instance or not</param>
        /// <returns></returns>
        public PPT.Presentation CreatePowerPointPresentation(PPT.Application powerPointApplication, bool showPowerPoint)
        {
            if (showPowerPoint)
            {
                return powerPointApplication.Presentations.Add(oTrue);
            }

            return powerPointApplication.Presentations.Add(oFalse);
        }

        /// <summary>
        /// Find a slide using its slide index ID
        /// </summary>
        /// <param name="presentation">Handle to a PPT.Presentation object to search through</param>
        /// <param name="slideId">SlideID to search for</param>
        /// <returns></returns>
        public PPT.Slide FindSlideByItsID(PPT.Presentation presentation, int slideId)
        {
            return presentation.Slides.FindBySlideID(slideId);
        }

        /// <summary>
        /// Open an existing PowerPoint presentation
        /// </summary>
        /// <param name="powerPointApplication">An instance of a PPT.Application object</param>
        /// <param name="pathAndFileName">Path (including filename) to the PowerPoint presentation</param>
        /// <returns>An instance of a PPT.Presentation object</returns>
        public PPT.Presentation OpenExistingPowerPointPresentation(
                PPT.Application powerPointApplication,
                string pathAndFileName)
        {
            return powerPointApplication.Presentations.Open(pathAndFileName, oFalse, oFalse, oFalse);
        }

        /// <summary>
        /// Save a PowerPoint presentation
        /// </summary>
        /// <param name="presentationToSave">Handle to PPT.Presentation object to save</param>
        /// <param name="pathAndFileName">Path (including filename) of where to save the presentation</param>
        /// <param name="fileType">PPT.PpSaveAsFileType object</param>
        /// <param name="embedTrueTypeFonts">Whether to embed TrueType fonts</param>
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
    }
}
namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;
    using Codesharper.PowerPoint.Helper.Enumerations;

    #endregion

    public interface ISlide
    {

        PPT.Slide AddSlideAtEndOfPresentation(PPT.Presentation presentationToAddSlideTo);

        int GetSlideCountInPresentation(PPT.Presentation presentation);

        PPT.Slide InsertSlideIntoPresentation(PPT.Presentation presentationToAddSlideTo, int indexOfSlide);

        PPT.SlideRange Clone(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex = 0);
    }
}

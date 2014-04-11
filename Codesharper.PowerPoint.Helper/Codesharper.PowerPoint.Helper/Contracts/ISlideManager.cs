namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using System.Collections.Generic;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;
    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Objects;

    #endregion

    public interface ISlideManager
    {

        PPT.Slide AddSlideToEnd(PPT.Presentation presentationToAddSlideTo);

        PPT.Slide AddSlideToStart(PPT.Presentation presentationToAddSlideTo);

        int GetSlideCount(PPT.Presentation presentation);

        PPT.Slide InsertSlide(PPT.Presentation presentationToAddSlideTo, int indexOfSlide);

        PPT.SlideRange CloneSlide(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex = 0);

        PPT.Slide SlideLayout(PPT.Slide slide, PPT.PpSlideLayout slideLayout);

        PPT.Slide MoveSlide(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex=0);

        void DeleteSlide(PPT.Slide slide);

        PPT.Slide AddComment(PPT.Slide slide, SlideComment comment);

        PPT.Slide DeleteComment(PPT.Slide slide, SlideComment comment);

        int CountComments(PPT.Slide slide);

        List<SlideComment> GetSlideComments(PPT.Slide slide);
    }
}

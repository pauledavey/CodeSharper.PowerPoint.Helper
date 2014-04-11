using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Implementations
{
    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Enumerations;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    public class SlideManager : ISlide
    {
        public PPT.Slide AddSlideAtEndOfPresentation(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(
                    (presentationToAddSlideTo.Slides.Count + 1),
                    PPT.PpSlideLayout.ppLayoutBlank);
        }

        public int GetSlideCountInPresentation(PPT.Presentation presentation)
        {
            return presentation.Slides.Count;
        }

        public PPT.Slide InsertSlideIntoPresentation(PPT.Presentation presentationToAddSlideTo, int indexOfSlide)
        {
            return presentationToAddSlideTo.Slides.Add(indexOfSlide, PPT.PpSlideLayout.ppLayoutBlank);
        }

       
        public PPT.SlideRange Clone(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex = 0)
        {
            var dupeSlide = slide.Duplicate();

            switch (destination)
            {
                    case Locations.Location.First:
                        dupeSlide.MoveTo(1);
                        break;

                    case Locations.Location.Last:
                        dupeSlide.MoveTo((presentation.Slides.Count));
                        break;

                    case Locations.Location.Custom:
                        dupeSlide.MoveTo(locationIndex);
                        break;
            }

            return dupeSlide;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Implementations
{
    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Objects;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    public class SlideManager : ISlideManager
    {
        public PPT.Slide AddSlideToEnd(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(
                    (presentationToAddSlideTo.Slides.Count + 1),
                    PPT.PpSlideLayout.ppLayoutBlank);
        }

        public PPT.Slide AddSlideToStart(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(1, PPT.PpSlideLayout.ppLayoutBlank);
        }

        public int GetSlideCount(PPT.Presentation presentation)
        {
            return presentation.Slides.Count;
        }

        public PPT.Slide InsertSlide(PPT.Presentation presentationToAddSlideTo, int indexOfSlide)
        {
            return presentationToAddSlideTo.Slides.Add(indexOfSlide, PPT.PpSlideLayout.ppLayoutBlank);
        }

       
        public PPT.SlideRange CloneSlide(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex = 0)
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

        public PPT.Slide SlideLayout(PPT.Slide slide, PPT.PpSlideLayout slideLayout)
        {
            slide.Layout = slideLayout;
            return slide;
        }

        public PPT.Slide MoveSlide(PPT.Presentation presentation, PPT.Slide slide, Locations.Location destination, int locationIndex = 0)
        {
            switch (destination)
            {
                case Locations.Location.First:
                    slide.MoveTo(1);
                    break;

                case Locations.Location.Last:
                    slide.MoveTo((presentation.Slides.Count));
                    break;

                case Locations.Location.Custom:
                    slide.MoveTo(locationIndex);
                    break;
            }

            return slide;
        }

        public void DeleteSlide(PPT.Slide slide)
        {
            slide.Delete();
        }

        public PPT.Slide AddComment(PPT.Slide slide, SlideComment comment)
        {
            PPT.Comment newComment = slide.Comments.Add(
                    comment.LeftPosition,
                    comment.TopPosition,
                    comment.Author,
                    comment.AuthorInitials,
                    comment.Comment);

            return slide;
        }

        public PPT.Slide DeleteComment(PPT.Slide slide, SlideComment slideComment)
        {
            foreach (PPT.Comment comment in slide.Comments)
            {
                if (comment.Text == slideComment.Comment && comment.Author == slideComment.Author && comment.AuthorInitials == slideComment.AuthorInitials)
                {
                    comment.Delete();
                }
            }

            return slide;
        }

        public int CountComments(PPT.Slide slide)
        {
            return slide.Comments.Count;
        }

        public List<SlideComment> GetSlideComments(PPT.Slide slide)
        {
            return (from PPT.Comment comment in slide.Comments select new SlideComment()
                                                                          {
                                                                                  Author = comment.Author,
                                                                                  AuthorInitials = comment.AuthorInitials,
                                                                                  Comment = comment.Text,
                                                                                  LeftPosition = comment.Left, 
                                                                                  TopPosition = comment.Top
                                                                          }).ToList();
        }
    }
}

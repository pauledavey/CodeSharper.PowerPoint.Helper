namespace Codesharper.PowerPoint.Helper.Implementations
{
    using System.Collections.Generic;
    using System.Linq;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Objects;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    public class SlideManager : ISlideManager
    {
        /// <summary>
        ///     Add a comment to the specified slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance to add the comment to</param>
        /// <param name="comment">SlideComment object containing configuration for the slide comment</param>
        /// <returns></returns>
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

        /// <summary>
        ///     Add a slide to the end of a presentation
        /// </summary>
        /// <param name="presentationToAddSlideTo">PPT.Presentation object to add slide to</param>
        /// <returns>PPT.Slide object instance</returns>
        public PPT.Slide AddSlideToEnd(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(
                    (presentationToAddSlideTo.Slides.Count + 1),
                    PPT.PpSlideLayout.ppLayoutBlank);
        }

        /// <summary>
        ///     Add Slide to start of a presentation
        /// </summary>
        /// <param name="presentationToAddSlideTo">PPT.Presentation object to add slide to</param>
        /// <returns>PPT.Slide object instance</returns>
        public PPT.Slide AddSlideToStart(PPT.Presentation presentationToAddSlideTo)
        {
            return presentationToAddSlideTo.Slides.Add(1, PPT.PpSlideLayout.ppLayoutBlank);
        }

        /// <summary>
        ///     Clone an existing slide (make a copy)
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <param name="slide">PPT.Slide instance that is to be cloned</param>
        /// <param name="destination">Destination for the cloned slide</param>
        /// <param name="locationIndex">Optional index for the new slide (slide.Index)</param>
        /// <returns></returns>
        public PPT.SlideRange CloneSlide(
                PPT.Presentation presentation,
                PPT.Slide slide,
                Locations.Location destination,
                int locationIndex = 0)
        {
            PPT.SlideRange dupeSlide = slide.Duplicate();

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

        /// <summary>
        ///     Returns the number of comments attached to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <returns>Count of comments</returns>
        public int CountComments(PPT.Slide slide)
        {
            return slide.Comments.Count;
        }

        /// <summary>
        ///     Delete the specified comment from a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance to delete the comment from</param>
        /// <param name="slideComment">The SlideComment object to delete</param>
        /// <returns></returns>
        public PPT.Slide DeleteComment(PPT.Slide slide, SlideComment slideComment)
        {
            foreach (PPT.Comment comment in slide.Comments)
            {
                if (comment.Text == slideComment.Comment && comment.Author == slideComment.Author
                    && comment.AuthorInitials == slideComment.AuthorInitials)
                {
                    comment.Delete();
                }
            }

            return slide;
        }

        /// <summary>
        ///     Delete the specified slide from the presentation
        /// </summary>
        /// <param name="slide">PPT.Slide object instance to delete</param>
        public void DeleteSlide(PPT.Slide slide)
        {
            slide.Delete();
        }

        /// <summary>
        ///     Export a slide to the specified image format
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="filePathAndName">File and Path name for the export target</param>
        /// <param name="exportFormat">Format to export to</param>
        /// <param name="scaleWidth">Scale with of the image</param>
        /// <param name="scaleHeight">Scale height of the image</param>
        public void Export(
                PPT.Slide slide,
                string filePathAndName,
                ImageFormats.Formats exportFormat,
                int scaleWidth = 1280,
                int scaleHeight = 1024)
        {
            slide.Export(filePathAndName, exportFormat.ToString(), scaleWidth, scaleHeight);
        }

        /// <summary>
        ///     Export all slides to the specified file formats
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <param name="filePath">File Path to export all the slides to (excluding filename(s))</param>
        /// <param name="exportFormat">Format to export to</param>
        /// <param name="scaleWidth">Scale with of the image</param>
        /// <param name="scaleHeight">Scale height of the image</param>
        public void ExportAll(
                PPT.Presentation presentation,
                string filePath,
                ImageFormats.Formats exportFormat,
                int scaleWidth = 1280,
                int scaleHeight = 1024)
        {
            foreach (PPT.Slide slide in presentation.Slides)
            {
                slide.Export(
                        filePath + @"slide" + slide.SlideIndex + "." + exportFormat,
                        exportFormat.ToString(),
                        scaleWidth,
                        scaleHeight);
            }
        }

        /// <summary>
        ///     Gets the slide comments
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <returns>A list of SlideComments from the slide</returns>
        public List<SlideComment> GetSlideComments(PPT.Slide slide)
        {
            return (from PPT.Comment comment in slide.Comments
                    select
                            new SlideComment
                                {
                                        Author = comment.Author,
                                        AuthorInitials = comment.AuthorInitials,
                                        Comment = comment.Text,
                                        LeftPosition = comment.Left,
                                        TopPosition = comment.Top
                                }).ToList();
        }

        /// <summary>
        ///     Get the total number of slides in the presentation
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <returns>Total number of slides in the presentation</returns>
        public int GetSlideCount(PPT.Presentation presentation)
        {
            return presentation.Slides.Count;
        }

        /// <summary>
        ///     Inserts a slide at the given index
        /// </summary>
        /// <param name="presentationToAddSlideTo">PPT.Presentation object to add slide to</param>
        /// <param name="indexOfSlide">Index for the new slide</param>
        /// <returns></returns>
        public PPT.Slide InsertSlide(PPT.Presentation presentationToAddSlideTo, int indexOfSlide)
        {
            return presentationToAddSlideTo.Slides.Add(indexOfSlide, PPT.PpSlideLayout.ppLayoutBlank);
        }

        /// <summary>
        ///     Moves a slide, changing its position with the presentation
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <param name="slide">PPT.Slide to move</param>
        /// <param name="destination">Destination location for the slide</param>
        /// <param name="locationIndex">Optional Index for the slides destination</param>
        /// <returns></returns>
        public PPT.Slide MoveSlide(
                PPT.Presentation presentation,
                PPT.Slide slide,
                Locations.Location destination,
                int locationIndex = 0)
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

        /// <summary>
        /// Retuns the dimensions of the slide (width and height)
        /// </summary>
        /// <param name="presentation">The presentation object (all slides are the same size)</param>
        /// <returns>a SlideDimensions object containing the height and width a slide will have in the presentation</returns>
        public SlideDimensions GetSlideDimensions(PPT.Presentation presentation)
        {
            return new SlideDimensions() { slideHeight = presentation.PageSetup.SlideHeight, 
                                           slideWidth = presentation.PageSetup.SlideWidth
                                         };
        }


        /// <summary>
        ///     Configure slide transiton effect for when a slide is loading
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="effect">The type of effect to apply</param>
        /// <param name="speed">Speed of the effect</param>
        /// <returns></returns>
        public void SetSlideTransition(PPT.Slide slide, PPT.PpEntryEffect effect, PPT.PpTransitionSpeed speed)
        {
            slide.SlideShowTransition.EntryEffect = effect;
            slide.SlideShowTransition.Speed = speed;
        }

        /// <summary>
        ///     Set the slide layout using one of PowerPoints builtin in layout templates
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="slideLayout">PPT.PpSlideLayout to apply</param>
        /// <returns></returns>
        public PPT.Slide SlideLayout(PPT.Slide slide, PPT.PpSlideLayout slideLayout)
        {
            slide.Layout = slideLayout;
            return slide;
        }
    }
}
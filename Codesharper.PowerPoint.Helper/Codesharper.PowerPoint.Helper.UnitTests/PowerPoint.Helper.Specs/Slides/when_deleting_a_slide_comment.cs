namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;
    using Codesharper.PowerPoint.Helper.Objects;

    using Microsoft.Office.Interop.PowerPoint;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_deleting_a_slide_comment : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private int initialCommentCount;

        private SlideManager slideManager;

        private SlideComment slideComment;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideComment = new SlideComment
                                    {
                                            Author = "Test User",
                                            AuthorInitials = "TA",
                                            Comment = "This is a test comment",
                                            LeftPosition = 100f,
                                            TopPosition = 100f
                                    };

            this.slideManager.AddComment(this.slideHandle, slideComment);

            this.initialCommentCount = this.slideManager.CountComments(this.slideHandle);
        }

        protected override void When()
        {
               this.slideHandle = this.slideManager.DeleteComment(this.slideHandle, this.slideComment); 
        }

        

        [Test]
        public void then_total_number_of_comments_should_decrease_by_one()
        {
            this.slideManager.CountComments(this.slideHandle).ShouldBeLessThan(this.initialCommentCount);

        }
    }
}

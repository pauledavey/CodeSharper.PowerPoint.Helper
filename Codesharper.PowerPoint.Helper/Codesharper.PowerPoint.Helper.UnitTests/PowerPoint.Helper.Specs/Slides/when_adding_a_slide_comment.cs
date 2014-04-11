namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;
    using Codesharper.PowerPoint.Helper.Objects;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_adding_a_slide_comment : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private int initialCommentCount;

        private SlideManager slideManager;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.initialCommentCount = this.slideManager.CountComments(this.slideHandle);
        }

        protected override void When()
        {
            this.slideManager.AddComment(
                    this.slideHandle,
                    new SlideComment()
                        {
                                Author = "Test Author",
                                AuthorInitials = "TA",
                                Comment = "This is a test comment",
                                LeftPosition = 100f,
                                TopPosition = 100f
                        });
        }

        

        [Test]
        public void then_total_number_of_comments_should_increase_by_one()
        {
            this.slideManager.CountComments(this.slideHandle).ShouldBeGreaterThan(this.initialCommentCount);

        }
    }
}

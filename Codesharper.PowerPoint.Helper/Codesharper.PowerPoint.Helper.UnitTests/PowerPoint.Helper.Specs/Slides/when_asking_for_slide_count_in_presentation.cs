namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_asking_for_slide_count_in_presentation : SpecsFor<PresentationManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private int slideCount;

        private SlideManager slideManager;
        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle, false);
        }

        protected override void When()
        {
            this.slideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        [Test]
        public void then_it_should_return_an_integer()
        {
            this.slideCount.ShouldBeType<int>();
        }
    }
}

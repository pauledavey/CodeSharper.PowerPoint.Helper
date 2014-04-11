namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_adding_a_slide_at_the_start_of_presentation : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private int initialSlideCount;

        private SlideManager slideManager;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            this.initialSlideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        protected override void When()
        {
            this.slideHandle = this.slideManager.AddSlideToStart(this.presentationHandle);
        }

        [Test]
        public void then_it_should_not_error()
        {
            this.slideHandle.ShouldNotBeNull();
        }


        [Test]
        public void then_total_number_of_slides_should_increase_by_one()
        {
            this.slideManager.GetSlideCount(this.presentationHandle).ShouldBeGreaterThan(this.initialSlideCount);

        }
    }
}

namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_deleting_a_slide : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slide;

        private Microsoft.Office.Interop.PowerPoint.Slide slide2;

        private int initialSlideCount;

        private SlideManager slideManager;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            slide = this.slideManager.AddSlideToEnd(this.presentationHandle);
            slide2 = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.initialSlideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        protected override void When()
        {
           this.slideManager.DeleteSlide(slide2);
        }

        [Test]
        public void then_total_number_of_slides_should_decrease_by_one()
        {
            this.slideManager.GetSlideCount(this.presentationHandle).ShouldBeLessThan(this.initialSlideCount);

        }
    }
}

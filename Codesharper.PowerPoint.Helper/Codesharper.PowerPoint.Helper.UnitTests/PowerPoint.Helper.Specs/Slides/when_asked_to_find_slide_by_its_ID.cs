namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    public class when_asked_to_find_slide_by_its_ID : SpecsFor<SlideManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;
        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;
        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;
        private Microsoft.Office.Interop.PowerPoint.Slide returnedSlide;
        private SlideManager slideManager;
        private PresentationManager presentation = new PresentationManager();
        private int slideID;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.presentation.CreatePowerPointPresentation(this.powerpointHandle,false);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideID = this.slideHandle.SlideID;
        }

        protected override void When()
        {
            this.returnedSlide = this.SUT.FindSlideByItsID(this.presentationHandle, this.slideID);

        }

        [Test]
        public void then_a_slide_should_be_returned()
        {
            this.returnedSlide.ShouldNotBeNull();
        }
    }
}

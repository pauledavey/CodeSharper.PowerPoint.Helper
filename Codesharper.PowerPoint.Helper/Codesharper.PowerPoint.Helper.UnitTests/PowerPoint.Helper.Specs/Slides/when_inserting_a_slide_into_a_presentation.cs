namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_inserting_a_slide_into_a_presentation : SpecsFor<PresentationManager>
    {
        private readonly PowerPointApplicationManager applicationHandler = new PowerPointApplicationManager();

        private readonly PresentationManager presentationHandler = new PresentationManager();

        private SlideManager slideManager;

        private int indexOfSlide = 1;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentation;

        private Microsoft.Office.Interop.PowerPoint.Slide slide;

        [Test]
        public void then_it_should_not_error()
        {
            this.slide.ShouldNotBeNull();
        }

        [Test]
        public void then_there_should_be_more_than_0_slides_in_the_presentation()
        {
            this.presentation.Slides.Count.ShouldBeGreaterThan(0);
        }

        protected override void Given()
        {
            var pptApplication = this.applicationHandler.CreatePowerPointApplication();
            this.slideManager = new SlideManager();
            this.presentation = this.presentationHandler.CreatePowerPointPresentation(pptApplication,false);
        }

        protected override void When()
        {
            this.slide = this.slideManager.InsertSlide(this.presentation, this.indexOfSlide);
        }
    }
}
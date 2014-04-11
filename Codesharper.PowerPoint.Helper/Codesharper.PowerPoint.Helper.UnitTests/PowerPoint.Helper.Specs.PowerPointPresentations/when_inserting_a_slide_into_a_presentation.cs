namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class when_inserting_a_slide_into_a_presentation : SpecsFor<Presentation>
    {
        private readonly PowerPointApplication applicationHandler = new PowerPointApplication();

        private readonly Presentation presentationHandler = new Presentation();

        private SlideManager slideManager;

        private int indexOfSlide = 1;

        private PPT.Presentation presentation;

        private PPT.Slide slide;

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
            this.slide = this.slideManager.InsertSlideIntoPresentation(this.presentation, this.indexOfSlide);
        }
    }
}
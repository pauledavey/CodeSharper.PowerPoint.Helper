namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using System.ComponentModel;

    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Interop.PowerPoint;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_cloning_a_slide_in_a_presentation : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;


        private Slide slide;

        private int initialSlideCount;

        private SlideManager slideManager;

        private SlideRange slideRange;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            slide = this.slideManager.AddSlideToEnd(presentationHandle);
            this.initialSlideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        protected override void When()
        {
            slideRange = this.slideManager.CloneSlide(this.presentationHandle, slide, Locations.Location.Last);
        }

        [Test]
        public void then_it_should_not_error()
        {
            this.slideRange.ShouldNotBeNull();
        }


        [Test]
        public void then_total_number_of_slides_should_increase_by_one()
        {
            this.slideManager.GetSlideCount(this.presentationHandle).ShouldBeGreaterThan(this.initialSlideCount);
        }
    }
}

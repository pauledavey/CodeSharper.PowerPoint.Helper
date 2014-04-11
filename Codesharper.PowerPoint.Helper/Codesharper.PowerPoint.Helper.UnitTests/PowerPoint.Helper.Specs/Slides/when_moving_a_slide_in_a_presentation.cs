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

    public class when_moving_a_slide_in_a_presentation : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Slide slide;

        private SlideManager slideManager;

        private Slide slide2;

        private Slide slide3;

        private const string slideName = "TestSlide";
        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            slide = this.slideManager.AddSlideToEnd(presentationHandle);
            slide2 = this.slideManager.AddSlideToEnd(presentationHandle);
            slide2.Name = slideName;
        }

        protected override void When()
        {
            this.slide3 = this.slideManager.MoveSlide(this.presentationHandle, slide2, Locations.Location.First);
        }

        
        [Test]
        public void then_the_slide_should_have_moved_from_the_end_position_to_the_first_position()
        {
            this.presentationHandle.Slides[1].Name.ShouldEqual(slideName);
            
        }
    }
}

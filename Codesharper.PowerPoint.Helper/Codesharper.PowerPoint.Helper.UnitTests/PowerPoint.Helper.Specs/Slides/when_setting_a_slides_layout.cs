namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using System.ComponentModel;

    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Interop.PowerPoint;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_setting_a_slides_layout : SpecsFor<PresentationManager>
    {
        private readonly PowerPointApplicationManager applicationHandler = new PowerPointApplicationManager();

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private SlideManager slideManager;

        private PpSlideLayout slideLayout;
        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = this.applicationHandler.CreatePowerPointApplication();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(powerpointHandle,false);
            this.slideManager.AddSlideToEnd(this.presentationHandle);
            slideLayout = this.presentationHandle.Slides[1].Layout;
        }

        protected override void When()
        {
            this.presentationHandle.Slides[1].Layout = PpSlideLayout.ppLayoutTwoObjects;
        }

        [Test]
        public void then_the_slide_layout_should_be_set_to_the_new_layout()
        {
            this.presentationHandle.Slides[1].Layout.ShouldNotEqual(this.slideLayout);
        }
    }
}

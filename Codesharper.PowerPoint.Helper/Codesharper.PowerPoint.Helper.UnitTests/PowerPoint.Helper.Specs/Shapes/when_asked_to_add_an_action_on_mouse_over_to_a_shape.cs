namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using Shapes = Codesharper.PowerPoint.Helper.Implementations.ShapesManager;

    public class when_asked_to_add_an_action_on_mouse_over_to_a_shape : SpecsFor<PresentationManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private Shapes shapesHandler;

        private Microsoft.Office.Interop.PowerPoint.Shape returnedShape;

        private SlideManager slideManager;
        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle, false);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.shapesHandler = new Shapes();
            this.returnedShape = this.shapesHandler.AddTextBoxToSlide(this.slideHandle, MsoTextOrientation.msoTextOrientationHorizontal, 100f, 100f, 100f, 100f);
        }

        protected override void When()
        {
          this.shapesHandler.AddMouseOverActionToShape(this.returnedShape, PpActionType.ppActionNextSlide);

        }

        [Test]
        public void then_the_shape_should_not_be_null()
        {
            this.returnedShape.ShouldNotBeNull();
        }

        [Test]
        public void then_the_shape_should_have_an_attached()
        {
            this.returnedShape.ActionSettings[PpMouseActivation.ppMouseOver].Action.ShouldEqual(
                    PpActionType.ppActionNextSlide);
        }
    }
}

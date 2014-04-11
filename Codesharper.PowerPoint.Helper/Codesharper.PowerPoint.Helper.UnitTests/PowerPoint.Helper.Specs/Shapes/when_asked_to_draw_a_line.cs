namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    public class when_asked_to_draw_a_line : SpecsFor<PresentationManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private ShapesManager shapesHandler;

        private SlideManager slideManager;

        private Microsoft.Office.Interop.PowerPoint.Shape returnedShape;

        protected override void Given()
        {
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.slideManager = new SlideManager();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle, false);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.shapesHandler = new ShapesManager();
        }

        protected override void When()
        {
            this.returnedShape = this.shapesHandler.DrawLine(this.slideHandle, 100f, 200f, 10f, 50f);
        }

        [Test]
        public void then_the_shape_is_not_null()
        {
            this.returnedShape.ShouldNotBeNull();
        }
    }
}

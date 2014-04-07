namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    using SpecsFor;

    public class when_asked_to_add_a_table_to_a_slide : SpecsFor<Presentation>
    {
        private PPT.Application powerpointHandle;

        private PPT.Presentation presentationHandle;

        private PPT.Slide slideHandle;

        private Shapes shapesHandler;

        private PPT.Shape returnedShape;

        protected override void Given()
        {
            this.powerpointHandle = new PPT.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(powerpointHandle, false);
            this.slideHandle = this.SUT.AddSlideAtEndOfPresentation(presentationHandle);
            this.shapesHandler = new Shapes();
        }

        protected override void When()
        {
            returnedShape = this.shapesHandler.AddTableToSlide(this.slideHandle, 4, 2, 100f, 100f, 100f, 100f);
        }

        [Test]
        public void then_the_shape_is_not_null()
        {
            returnedShape.ShouldNotBeNull();
        }

        [Test]
        public void then_there_should_be_a_shape_on_the_slide()
        {
            this.slideHandle.Shapes.Count.ShouldEqual(1);
        }
    }
}

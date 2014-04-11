namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using Shapes = Codesharper.PowerPoint.Helper.Implementations.ShapesManager;

    public class when_asked_to_set_a_textboxs_text : SpecsFor<PresentationManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private Shapes shapesHandler;

        private Microsoft.Office.Interop.PowerPoint.Shape returnedShape;

        private SlideManager slideManager;

        private const string textboxText = "Test Text";

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
            this.shapesHandler.SetTextBoxText(this.returnedShape, textboxText);
        }

        [Test]
        public void then_the_textbox_text_should_not_be_null()
        {
            this.returnedShape.TextEffect.Text.ShouldNotBeEmpty();
        }

        [Test]
        public void the_text_should_equal_out_set_value()
        {
            this.returnedShape.TextEffect.Text.ShouldEqual(textboxText);
        }
    }
}

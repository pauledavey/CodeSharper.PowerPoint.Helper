namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using System;

    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    public class when_asked_to_add_a_picture : SpecsFor<PresentationManager>
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
            var filePath = Environment.CurrentDirectory + "\\testpicture.png";

            this.returnedShape = this.shapesHandler.AddPicture(
                    this.slideHandle,
                    filePath,
                    100f,
                    100f,
                    300f,
                    300f);
        }

        [Test]
        public void then_the_shape_should_be_a_picture()
        {
            this.returnedShape.PictureFormat.Brightness.ShouldBeGreaterThan(0);
        }
    }
}

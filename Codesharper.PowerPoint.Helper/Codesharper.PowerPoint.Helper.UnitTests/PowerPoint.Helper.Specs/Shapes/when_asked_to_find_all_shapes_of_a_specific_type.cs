namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using System.Collections.Generic;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    public class when_asked_to_find_all_shapes_of_a_specific_type : SpecsFor<PresentationManager>
    {
        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private SlideManager slideManager;

        private ShapesManager shapesHandler = new ShapesManager();

        private List<ShapesofType> listOfShapes = new List<ShapesofType>();

        protected override void Given()
        {
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.slideManager = new SlideManager();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle, false);

            // Add some slides to the presentation
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);

            // Add some shapes to slides
            this.shapesHandler.DrawShape(
                    this.presentationHandle.Slides[1],
                    MsoAutoShapeType.msoShapeRectangle,
                    100f,
                    100f,
                    100f,
                    100f);

            this.shapesHandler.DrawShape(
                    this.presentationHandle.Slides[2],
                    MsoAutoShapeType.msoShapeRectangle,
                    100f,
                    100f,
                    100f,
                    100f);

        }

        protected override void When()
        {
            this.listOfShapes = this.shapesHandler.FindShapesInPresentation(presentationHandle, MsoAutoShapeType.msoShapeRectangle);

        }

        [Test]
        public void then_the_list_of_shapes_should_contain_two_objects_as_we_have_two_rectangles()
        {
            this.listOfShapes.Count.ShouldEqual(2);
        }
    }
}

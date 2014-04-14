namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using System;

    using Codesharper.PowerPoint.Helper.Enumerations;
    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_exporting_a_slide : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private int initialSlideCount;

        private SlideManager slideManager;

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
            this.initialSlideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        protected override void When()
        {
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideManager.Export(this.slideHandle, Environment.CurrentDirectory + @"\testslide.png", ImageFormats.Formats.png);
        }

        
        [Test]
        public void then_the_slide_should_export_as_specified()
        {
            bool answer = System.IO.File.Exists(Environment.CurrentDirectory + @"\testslide.png");
            answer.ShouldBeTrue();

        }
    }
}

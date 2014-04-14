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

    public class when_exporting_all_slides : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        private Microsoft.Office.Interop.PowerPoint.Slide slideHandle;

        private int slideCount;

        private SlideManager slideManager;

        private int counter;

        private string filePath = Environment.CurrentDirectory + @"\when_exporting_all_slides_test\";

        protected override void Given()
        {
            this.slideManager = new SlideManager();
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
        }

        protected override void When()
        {
            if (!System.IO.Directory.Exists(filePath))
            {
                System.IO.Directory.CreateDirectory(filePath);
            }

            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);
            this.slideManager.ExportAll(this.presentationHandle, filePath, ImageFormats.Formats.png);
            this.slideCount = this.slideManager.GetSlideCount(this.presentationHandle);
        }

        
        [Test]
        public void then_the_number_of_exported_slides_should_equal_the_number_of_slides_in_the_presentation()
        {

            var folder =
                    System.IO.Directory.EnumerateFiles(filePath);

            foreach (string file in System.IO.Directory.EnumerateFiles(filePath))
            {
                if (file.Contains(@"slide") && file.Contains(".png"))
                {
                    counter++;
                }
            }

            counter.ShouldEqual(this.slideCount);

        }
    }
}

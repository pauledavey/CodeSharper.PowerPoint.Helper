namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class when_asking_for_slide_count_in_presentation : SpecsFor<Presentation>
    {
        private PPT.Application powerpointHandle;

        private PPT.Presentation presentationHandle;

        private int slideCount;

        protected override void Given()
        {
            this.powerpointHandle = new PPT.Application();
            this.presentationHandle = SUT.CreatePowerPointPresentation(powerpointHandle, false);
        }

        protected override void When()
        {
            slideCount = this.SUT.GetSlideCountInPresentation(this.presentationHandle);
        }

        [Test]
        public void then_it_should_return_an_integer()
        {
            slideCount.ShouldBeType<int>();
        }
    }
}

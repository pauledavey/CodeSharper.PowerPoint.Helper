namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class when_adding_a_slide_at_the_end_of_presentation : SpecsFor<Presentation>
    {

        private PPT.Application powerpointHandle;

        private PPT.Presentation presentationHandle;

        private PPT.Slide slideHandle;

        private int initialSlideCount;

        protected override void Given()
        {
            this.powerpointHandle = new PPT.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(powerpointHandle,false);
            initialSlideCount = this.SUT.GetSlideCountInPresentation(presentationHandle);
        }

        protected override void When()
        {
            slideHandle = this.SUT.AddSlideAtEndOfPresentation(this.presentationHandle);
        }

        [Test]
        public void then_it_should_not_error()
        {
            this.slideHandle.ShouldNotBeNull();
        }


        [Test]
        public void then_total_number_of_slides_should_increase_by_one()
        {
            this.SUT.GetSlideCountInPresentation(this.presentationHandle).ShouldBeGreaterThan(initialSlideCount);

        }
    }
}

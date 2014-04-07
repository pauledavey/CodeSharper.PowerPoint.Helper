namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class when_creating_a_new_presentation : SpecsFor<Presentation>
    {

        private PPT.Application powerpointHandle;

        private PPT.Presentation presentationHandle;

        protected override void Given()
        {
            this.powerpointHandle = new PPT.Application();
        }

        protected override void When()
        {
            presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
        }

        [Test]
        public void then_it_should_not_error()
        {
            presentationHandle.ShouldNotBeNull();
        }
    }
}

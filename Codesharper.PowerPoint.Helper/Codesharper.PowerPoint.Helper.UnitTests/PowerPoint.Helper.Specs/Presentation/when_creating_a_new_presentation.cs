namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

    public class when_creating_a_new_presentation : SpecsFor<PresentationManager>
    {

        private Microsoft.Office.Interop.PowerPoint.Application powerpointHandle;

        private Microsoft.Office.Interop.PowerPoint.Presentation presentationHandle;

        protected override void Given()
        {
            this.powerpointHandle = new Microsoft.Office.Interop.PowerPoint.Application();
        }

        protected override void When()
        {
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(this.powerpointHandle,false);
        }

        [Test]
        public void then_it_should_not_error()
        {
            this.presentationHandle.ShouldNotBeNull();
        }
    }
}

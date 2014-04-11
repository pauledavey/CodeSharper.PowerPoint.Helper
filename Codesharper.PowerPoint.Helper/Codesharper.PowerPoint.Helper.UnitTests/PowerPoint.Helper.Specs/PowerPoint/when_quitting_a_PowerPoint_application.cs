namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

        public class when_quitting_a_PowerPoint_application : SpecsFor<PowerPointApplicationManager>
        {
            private Microsoft.Office.Interop.PowerPoint.Application appHandle;

            [Test]
            public void then_it_should_not_error()
            {
                this.SUT.ShouldNotBeNull();
            }

            protected override void Given()
            {
                this.appHandle = this.SUT.CreatePowerPointApplication();
            }

            protected override void When()
            {
                this.SUT.ClosePowerPointApplication(this.appHandle);
            }
        }
}
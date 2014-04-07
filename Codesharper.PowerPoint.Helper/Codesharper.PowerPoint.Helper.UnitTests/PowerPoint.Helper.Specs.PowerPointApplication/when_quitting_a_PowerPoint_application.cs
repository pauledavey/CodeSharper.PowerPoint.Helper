namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointApplication
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

        public class when_quitting_a_PowerPoint_application : SpecsFor<PowerPointApplication>
        {
            private PPT.Application appHandle;

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
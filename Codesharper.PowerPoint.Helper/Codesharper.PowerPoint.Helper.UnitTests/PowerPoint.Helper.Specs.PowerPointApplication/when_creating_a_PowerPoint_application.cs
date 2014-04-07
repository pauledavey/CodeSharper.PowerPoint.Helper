namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointApplication
{
    #region Using Directives

    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    #endregion

        public class when_creating_a_PowerPoint_application : SpecsFor<PowerPointApplication>
        {
            private Microsoft.Office.Interop.PowerPoint.Application _result;

            [Test]
            public void then_it_returns_a_PowerPoint_Application_object()
            {
                this._result.ShouldNotBeNull();
            }

            protected override void When()
            {
                this._result = this.SUT.CreatePowerPointApplication();
            }
        }
}
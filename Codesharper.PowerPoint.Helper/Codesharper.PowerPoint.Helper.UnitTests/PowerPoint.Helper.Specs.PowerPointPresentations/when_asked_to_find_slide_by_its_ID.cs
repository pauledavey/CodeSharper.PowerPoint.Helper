using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs.PowerPointPresentations
{
    using Codesharper.PowerPoint.Helper.Implementations;

    using NUnit.Framework;

    using Should;

    using PPT = Microsoft.Office.Interop.PowerPoint;

    using SpecsFor;

    public class when_asked_to_find_slide_by_its_ID : SpecsFor<Presentation>
    {
        private PPT.Application powerpointHandle;
        private PPT.Slide slideHandle;
        private PPT.Presentation presentationHandle;
        private PPT.Slide returnedSlide;
        private SlideManager slideManager;
        private Presentation presentation = new Presentation();
        private int slideID;

        protected override void Given()
        {
            slideManager = new SlideManager();
            this.powerpointHandle = new PPT.Application();
            this.presentationHandle = this.SUT.CreatePowerPointPresentation(powerpointHandle,false);
            slideHandle = this.slideManager.AddSlideAtEndOfPresentation(this.presentationHandle);
            slideID = slideHandle.SlideID;
        }

        protected override void When()
        {
            returnedSlide = this.SUT.FindSlideByItsID(presentationHandle, slideID);

        }

        [Test]
        public void then_a_slide_should_be_returned()
        {
            returnedSlide.ShouldNotBeNull();
        }
    }
}

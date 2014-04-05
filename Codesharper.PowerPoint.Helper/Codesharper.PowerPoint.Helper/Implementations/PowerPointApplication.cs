namespace Codesharper.PowerPoint.Helper.Implementations
{
    using Codesharper.PowerPoint.Helper.Contracts;
    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;

    public class PowerPointApplication : IPowerPointApplication
    {
        public PPT.Application CreatePowerPointApplication()
        {
            var pptAppHandle = new PPT.Application();
            return pptAppHandle;
        }

        public void ClosePowerPointApplication(PPT.Application powerPointApplication)
        {
            powerPointApplication.Quit();
        }
    }
}

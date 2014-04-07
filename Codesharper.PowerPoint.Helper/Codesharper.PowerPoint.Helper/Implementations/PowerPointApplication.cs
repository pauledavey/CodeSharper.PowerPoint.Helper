namespace Codesharper.PowerPoint.Helper.Implementations
{
    #region Using Directives

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class PowerPointApplication : OFFICE.IPowerPointApplication
    {

        public void ClosePowerPointApplication(PPT.Application powerPointApplication)
        {
            powerPointApplication.Quit();
        }

        public PPT.Application CreatePowerPointApplication()
        {
            var pptAppHandle = new PPT.Application();
            return pptAppHandle;
        }
    }
}
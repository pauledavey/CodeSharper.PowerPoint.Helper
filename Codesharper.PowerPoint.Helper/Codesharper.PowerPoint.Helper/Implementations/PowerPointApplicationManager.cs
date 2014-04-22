namespace Codesharper.PowerPoint.Helper.Implementations
{
    #region Using Directives

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;

    #endregion

    public class PowerPointApplicationManager : OFFICE.IPowerPointManager
    {
        public PowerPointApplicationManager()
        {
            
        }


        /// <summary>
        /// Close the PowerPoint application instance
        /// </summary>
        /// <param name="powerPointApplication">PowerPoint instance</param>
        public void ClosePowerPointApplication(PPT.Application powerPointApplication)
        {
            powerPointApplication.Quit();
        }

        /// <summary>
        /// Create an instance of the PowerPoint application
        /// </summary>
        /// <returns>A PPT.Application object</returns>
        public PPT.Application CreatePowerPointApplication()
        {
            var pptAppHandle = new PPT.Application();
            return pptAppHandle;
        }
    }
}
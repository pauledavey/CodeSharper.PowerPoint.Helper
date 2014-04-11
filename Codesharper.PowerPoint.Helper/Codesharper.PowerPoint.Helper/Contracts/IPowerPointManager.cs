namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    #endregion

    public interface IPowerPointManager
    {
        void ClosePowerPointApplication(PPT.Application powerPointApplication);

        PPT.Application CreatePowerPointApplication();
    }
}
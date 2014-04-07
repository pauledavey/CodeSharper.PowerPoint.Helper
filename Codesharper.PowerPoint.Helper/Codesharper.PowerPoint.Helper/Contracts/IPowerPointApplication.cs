namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    #endregion

    public interface IPowerPointApplication
    {
        void ClosePowerPointApplication(PPT.Application powerPointApplication);

        PPT.Application CreatePowerPointApplication();
    }
}
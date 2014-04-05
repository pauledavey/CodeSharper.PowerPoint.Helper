namespace Codesharper.PowerPoint.Helper.Contracts
{
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;
    
    public interface IPowerPointApplication
    {

        PPT.Application CreatePowerPointApplication();

        void ClosePowerPointApplication(PPT.Application powerPointApplication);
    }
}

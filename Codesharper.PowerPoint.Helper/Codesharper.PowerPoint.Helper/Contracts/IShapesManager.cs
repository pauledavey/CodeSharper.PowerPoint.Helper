namespace Codesharper.PowerPoint.Helper.Contracts
{
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    using System.Security.Cryptography.X509Certificates;

    public interface IShapesManager
    {

        PPT.Shape AddTextBoxToSlide(
                PPT.Slide slide,
                OFFICE.MsoTextOrientation orientation,
                float widthLocation,
                float heightLocation,
                float x,
                float y);

        void SetTextBoxText(PPT.Shape textbox, string text);

        PPT.Shape AddTableToSlide(PPT.Slide slide, int numRows, int numColumns, float xLocation, float yLocation, float width, float height);

        PPT.Shape DrawLine(PPT.Slide slide, float xStartLocation, float xEndLocation, float yStartLocation, float yEndLocation);

        PPT.Shape DrawShape(
                PPT.Slide slide,
                OFFICE.MsoAutoShapeType shapeType,
                float leftPosition,
                float topPosition,
                float width,
                float height);

        PPT.Shape AddPicture(
                PPT.Slide slide,
                string file,
                float leftPosition,
                float topPosition,
                float width,
                float height);
    }
}

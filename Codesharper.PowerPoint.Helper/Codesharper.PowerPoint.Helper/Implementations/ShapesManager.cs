using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Implementations
{
    using Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;
    using Microsoft.Office.Core;

    public class ShapesManager : IShapesManager
    {
        public PPT.Shape AddTextBoxToSlide(PPT.Slide slide, MsoTextOrientation orientation, float xLocation, float yLocation, float width, float height)
        {
            var textbox = slide.Shapes.AddTextbox(orientation, xLocation, yLocation, width, height);
            return textbox;
        }

        public void SetTextBoxText(PPT.Shape textbox, string text)
        {
            textbox.TextEffect.Text = text;
        }

        public PPT.Shape AddTableToSlide(PPT.Slide slide, int numRows, int numColumns, float xLocation, float yLocation, float width, float height)
        {
            var table = slide.Shapes.AddTable(numRows, numColumns, xLocation, yLocation, width, height);
            return table;
        }

        public PPT.Shape DrawLine(PPT.Slide slide, float xStartLocation, float xEndLocation, float yStartLocation, float yEndLocation)
        {
            return slide.Shapes.AddLine(xStartLocation, yStartLocation, xEndLocation, yEndLocation);
        }

        public PPT.Shape DrawShape(PPT.Slide slide, MsoAutoShapeType shapeType, float leftPosition, float topPosition, float width, float height)
        {
            return slide.Shapes.AddShape(shapeType, leftPosition, topPosition, width, height);
        }

        public PPT.Shape AddPicture(PPT.Slide slide, string file, float leftPosition, float topPosition, float width, float height)
        {
            var shapeOut = slide.Shapes.AddPicture(
                    file,
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    leftPosition,
                    topPosition,
                    width,
                    height);

            return shapeOut;
        }
    }
}

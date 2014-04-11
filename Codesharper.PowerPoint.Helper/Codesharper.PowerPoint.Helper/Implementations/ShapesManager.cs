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
    }
}

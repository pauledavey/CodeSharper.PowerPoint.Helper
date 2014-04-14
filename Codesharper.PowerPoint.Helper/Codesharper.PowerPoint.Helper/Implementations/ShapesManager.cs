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
        /// <summary>
        /// Add a Textbox to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object to add a textbox to</param>
        /// <param name="orientation">Orientation of the textbox</param>
        /// <param name="xLocation">x Location</param>
        /// <param name="yLocation">y Location</param>
        /// <param name="width">Textbox width</param>
        /// <param name="height">Textbox height</param>
        /// <returns></returns>
        public PPT.Shape AddTextBoxToSlide(PPT.Slide slide, MsoTextOrientation orientation, float xLocation, float yLocation, float width, float height)
        {
            var textbox = slide.Shapes.AddTextbox(orientation, xLocation, yLocation, width, height);
            return textbox;
        }

        /// <summary>
        /// Set the text in the textbox
        /// </summary>
        /// <param name="textbox">PPT.Shape that is a textbox</param>
        /// <param name="text">Text</param>
        public void SetTextBoxText(PPT.Shape textbox, string text)
        {
            textbox.TextEffect.Text = text;
        }

        /// <summary>
        /// Add a table to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="numRows">Number of rows to create in the table</param>
        /// <param name="numColumns">Number of columns to create in the table</param>
        /// <param name="xLocation">x location</param>
        /// <param name="yLocation">y location</param>
        /// <param name="width">Table shape width</param>
        /// <param name="height">Table shape height</param>
        /// <returns></returns>
        public PPT.Shape AddTableToSlide(PPT.Slide slide, int numRows, int numColumns, float xLocation, float yLocation, float width, float height)
        {
            var table = slide.Shapes.AddTable(numRows, numColumns, xLocation, yLocation, width, height);
            return table;
        }

        /// <summary>
        /// Draws a line on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="xStartLocation">Starting x location</param>
        /// <param name="xEndLocation">Ending x location</param>
        /// <param name="yStartLocation">Starting y location</param>
        /// <param name="yEndLocation">Ending y location</param>
        /// <returns></returns>
        public PPT.Shape DrawLine(PPT.Slide slide, float xStartLocation, float xEndLocation, float yStartLocation, float yEndLocation)
        {
            return slide.Shapes.AddLine(xStartLocation, yStartLocation, xEndLocation, yEndLocation);
        }

        /// <summary>
        /// Draw a shape on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="shapeType">The shape type. Any shape from the MsoAutoShapeType can be specified</param>
        /// <param name="leftPosition">x position</param>
        /// <param name="topPosition">y position</param>
        /// <param name="width">Shape width</param>
        /// <param name="height">Shape height</param>
        /// <returns></returns>
        public PPT.Shape DrawShape(PPT.Slide slide, MsoAutoShapeType shapeType, float leftPosition, float topPosition, float width, float height)
        {
            return slide.Shapes.AddShape(shapeType, leftPosition, topPosition, width, height);
        }

        /// <summary>
        /// Add an existing picture to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="file">The file path and name of the image</param>
        /// <param name="leftPosition">x Location</param>
        /// <param name="topPosition">y Location</param>
        /// <param name="width">Width of the picture</param>
        /// <param name="height">Height of the picture</param>
        /// <returns></returns>
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

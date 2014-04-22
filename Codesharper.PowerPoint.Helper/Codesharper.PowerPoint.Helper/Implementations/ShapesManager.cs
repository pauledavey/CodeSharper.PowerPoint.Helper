namespace Codesharper.PowerPoint.Helper.Implementations
{
    using System.Collections.Generic;
    using System.Linq;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Objects;

    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;

    public class ShapesManager : IShapesManager
    {

        public ShapesManager()
        {
            
        }

        /// <summary>
        ///     Add an existing picture to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="file">The file path and name of the image</param>
        /// <param name="leftPosition">x Location</param>
        /// <param name="topPosition">y Location</param>
        /// <param name="width">Width of the picture</param>
        /// <param name="height">Height of the picture</param>
        /// <returns></returns>
        public PPT.Shape AddPicture(
                PPT.Slide slide,
                string file,
                float leftPosition,
                float topPosition,
                float width,
                float height)
        {
            PPT.Shape shapeOut = slide.Shapes.AddPicture(
                    file,
                    OFFICE.MsoTriState.msoFalse,
                    OFFICE.MsoTriState.msoTrue,
                    leftPosition,
                    topPosition,
                    width,
                    height);

            return shapeOut;
        }

        /// <summary>
        ///     Add a table to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="numRows">Number of rows to create in the table</param>
        /// <param name="numColumns">Number of columns to create in the table</param>
        /// <param name="xLocation">x location</param>
        /// <param name="yLocation">y location</param>
        /// <param name="width">Table shape width</param>
        /// <param name="height">Table shape height</param>
        /// <returns></returns>
        public PPT.Shape AddTableToSlide(
                PPT.Slide slide,
                int numRows,
                int numColumns,
                float xLocation,
                float yLocation,
                float width,
                float height)
        {
            PPT.Shape table = slide.Shapes.AddTable(numRows, numColumns, xLocation, yLocation, width, height);
            return table;
        }

        /// <summary>
        ///     Add a Textbox to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object to add a textbox to</param>
        /// <param name="orientation">Orientation of the textbox</param>
        /// <param name="xLocation">x Location</param>
        /// <param name="yLocation">y Location</param>
        /// <param name="width">Textbox width</param>
        /// <param name="height">Textbox height</param>
        /// <returns></returns>
        public PPT.Shape AddTextBoxToSlide(
                PPT.Slide slide,
                OFFICE.MsoTextOrientation orientation,
                float xLocation,
                float yLocation,
                float width,
                float height)
        {
            PPT.Shape textbox = slide.Shapes.AddTextbox(orientation, xLocation, yLocation, width, height);
            return textbox;
        }

        /// <summary>
        ///     Draws a line on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="xStartLocation">Starting x location</param>
        /// <param name="xEndLocation">Ending x location</param>
        /// <param name="yStartLocation">Starting y location</param>
        /// <param name="yEndLocation">Ending y location</param>
        /// <returns></returns>
        public PPT.Shape DrawLine(
                PPT.Slide slide,
                float xStartLocation,
                float xEndLocation,
                float yStartLocation,
                float yEndLocation)
        {
            return slide.Shapes.AddLine(xStartLocation, yStartLocation, xEndLocation, yEndLocation);
        }

        /// <summary>
        ///     Draw a shape on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="shapeType">The shape type. Any shape from the MsoAutoShapeType can be specified</param>
        /// <param name="leftPosition">x position</param>
        /// <param name="topPosition">y position</param>
        /// <param name="width">Shape width</param>
        /// <param name="height">Shape height</param>
        /// <returns></returns>
        public PPT.Shape DrawShape(
                PPT.Slide slide,
                OFFICE.MsoAutoShapeType shapeType,
                float leftPosition,
                float topPosition,
                float width,
                float height)
        {
            return slide.Shapes.AddShape(shapeType, leftPosition, topPosition, width, height);
        }

        /// <summary>
        ///     Find all shapes of a type in a presentation
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <param name="shapeType">The shape type to look for</param>
        /// <returns></returns>
        public List<ShapesofType> FindShapesInPresentation(PPT.Presentation presentation, OFFICE.MsoAutoShapeType shapeType)
        {
            return (from PPT.Slide slide in presentation.Slides
                    from PPT.Shape shape in slide.Shapes
                    where shape.AutoShapeType == shapeType
                    select new ShapesofType { shape = shape, shapeType = shape.Type, slide = slide }).ToList();
        }

        /// <summary>
        ///     Set the text in the textbox
        /// </summary>
        /// <param name="textbox">PPT.Shape that is a textbox</param>
        /// <param name="text">Text</param>
        public void SetTextBoxText(PPT.Shape textbox, string text)
        {
            textbox.TextEffect.Text = text;
        }
    }
}
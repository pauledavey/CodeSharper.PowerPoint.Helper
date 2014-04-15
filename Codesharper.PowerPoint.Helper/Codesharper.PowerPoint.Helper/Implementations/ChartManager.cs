
namespace Codesharper.PowerPoint.Helper.Implementations
{
    using Microsoft.Office.Core;

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;

    using Codesharper.PowerPoint.Helper.Contracts;

    public class ChartManager : IChartManager
    {
        //public PPT.Slide CreateChart(PPT.Slide slide, XlChartType chartType, float xLocation, float yLocation, float width, float height)
        //{
        //    return slide.Shapes.AddChart(chartType, xLocation, yLocation, width, height);
        //}





        public void AddChartTitle(PPT.Shape chart, string titleText)
        {
            chart.Title = titleText;
        }

    }
}

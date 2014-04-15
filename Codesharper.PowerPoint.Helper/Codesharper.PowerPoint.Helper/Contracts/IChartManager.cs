using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Contracts
{
    using Microsoft.Office.Core;
    using PPT = Microsoft.Office.Interop.PowerPoint;

    public interface IChartManager
    {
        //PPT.Slide CreateChart(
        //        PPT.Slide slide,
        //        XlChartType chartType,
        //        float xLocation,
        //        float yLocation,
        //        float width,
        //        float height);

        void AddChartTitle(PPT.Shape chart, string titleText);


    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Codesharper.PowerPoint.Helper.Contracts
{
    using Microsoft.Office.Core;

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using EXCEL = Microsoft.Office.Interop.Excel;

    public interface IChartManager
    {
      //  PPT.Chart CreateChart(PPT.Slide slide, ChartConfiguration chartConfiguration);

        void AddChartTitle(PPT.Shape chart, string titleText);


    }

    public class ChartConfiguration
    {
        public XlChartType chartType
        {
            get;
            set;
        }

        public float xLocation
        {
            get;
            set;
        }

        public float yLocation
        {
            get;
            set;
        }

        public float width
        {
            get;
            set;
        }

        public float height
        {
            get;
            set;
        }
    }
}

namespace Codesharper.PowerPoint.Helper.Contracts
{
    #region Using Directives

    using System.Collections.Generic;
    using Microsoft.Office.Core;
    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using EXCEL = Microsoft.Office.Interop.Excel;

    #endregion

    public interface IChartManager
    {
        void AddChartLegend(PPT.Chart chart, ChartLegend chartLegend);

        void AddChartTitle(PPT.Chart chart, ChartTitle chartTitle);

        PPT.Chart CreateChart(PPT.Slide slide, string[] xAxisPoints, List<ChartSeries> datasets);

        void AddSeriesToExistingChart(PPT.Chart chart, ChartSeries series);
    }

    public class ChartTitle
    {
        public bool bold
        {
            get;
            set;
        }

        public int fontSize
        {
            get;
            set;
        }

        public bool italic
        {
            get;
            set;
        }

        public string titleText
        {
            get;
            set;
        }

        public bool underline
        {
            get;
            set;
        }
    }

    public class ChartLegend
    {
        public bool bold
        {
            get;
            set;
        }

        public int fontSize
        {
            get;
            set;
        }

        public bool italic
        {
            get;
            set;
        }

        public bool underline
        {
            get;
            set;
        }
    }

    public class ChartSeries
    {
        public string name
        {
            get;
            set;
        }

        public string[] seriesData
        {
            get;
            set;
        }

        public XlChartType seriesType
        {
            get;
            set;
        }
    }

    public class ChartConfiguration
    {
        public XlChartType chartType
        {
            get;
            set;
        }

        public float height
        {
            get;
            set;
        }

        public float width
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
    }
}
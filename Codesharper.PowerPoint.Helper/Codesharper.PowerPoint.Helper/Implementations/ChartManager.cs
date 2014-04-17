namespace Codesharper.PowerPoint.Helper.Implementations
{
    #region Using Directives

    using System.Collections.Generic;
    using System.Linq;
    using Codesharper.PowerPoint.Helper.Contracts;
    using Microsoft.Office.Core;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using EXCEL = Microsoft.Office.Interop.Excel;

    #endregion

    public class ChartManager : IChartManager
    {

        public void AddChartLegend(PPT.Chart chart, ChartLegend chartLegend)
        {
            chart.HasLegend = true;
            chart.Legend.Font.Italic = chartLegend.italic;
            chart.Legend.Font.Bold = chartLegend.bold;
            chart.Legend.Font.Underline = chartLegend.underline;
            chart.Legend.Font.Size = chartLegend.fontSize;
            chart.Refresh();
        }

        public void AddChartTitle(PPT.Chart chart, ChartTitle chartTitle)
        {
            chart.HasTitle = true;
            chart.ChartTitle.Text = chartTitle.titleText;
            chart.ChartTitle.Font.Italic = chartTitle.italic;
            chart.ChartTitle.Font.Bold = chartTitle.bold;
            chart.ChartTitle.Font.Underline = chartTitle.underline;
            chart.ChartTitle.Font.Size = chartTitle.fontSize;
            chart.Refresh();
        }

        public PPT.Chart CreateChart(PPT.Slide slide, string[] xAxisPoints, List<ChartSeries> datasets)
        {
            var chart = slide.Shapes.AddChart(XlChartType.xlLine, 10f, 10f, 900f, 400f).Chart;

            var workbook = (EXCEL.Workbook)chart.ChartData.Workbook;
            workbook.Windows.Application.Visible = false;

            var dataSheet = (EXCEL.Worksheet)workbook.Worksheets[1];
            dataSheet.Cells.ClearContents();
            dataSheet.Cells.Clear();
            dataSheet.Calculate();

            var sc = (PPT.SeriesCollection)chart.SeriesCollection();

            do
            {
                var seriesToDelete = sc.Item(1);
                seriesToDelete.Delete();
                chart.Refresh();
            }
            while (sc.Count != 0);

            //Build out the X-Axis Data Categories
            for (var i = 1; i != (xAxisPoints.Count() + 1); i++)
            {
                dataSheet.Cells.Range["A" + i].Value2 = xAxisPoints[(i - 1)];
                chart.Refresh();
            }

            var intLetter = 1;
            var cellNumber = 1;

            for (var j = 0; j < datasets.Count; j++)
            {
                var letter = IntToLetters((intLetter + 1));

                // each one of these is a dataset.
                foreach (var value in datasets[j].seriesData)
                {
                    var cellPosition = letter + cellNumber.ToString();
                    dataSheet.Cells.Range[cellPosition].Value2 = value;
                    cellNumber++;
                    chart.Refresh();
                }

                // we have populate the sheet with new values, now we need to create a series for it!
                var columnsRange = dataSheet.UsedRange.Columns;
                var rowsRange = dataSheet.UsedRange.Rows;

                var columnCount = columnsRange.Columns.Count;
                var rowCount = rowsRange.Rows.Count;
                var lastColumnLetter = IntToLetters(columnCount);

                var newSeries = sc.NewSeries();
                newSeries.Name = datasets[j].name;
                newSeries.XValues = "'Sheet1'!$A$1:$A$" + rowCount;
                newSeries.Values = "'Sheet1'!$" + lastColumnLetter + "$1:$" + lastColumnLetter + "$" + rowCount;
                newSeries.ChartType = datasets[j].seriesType;

                intLetter++;
                cellNumber = 1;
                chart.Refresh();
            }

            chart.HasTitle = false;
            chart.HasLegend = false;
            chart.Refresh();

            return chart;
        }

        public void AddSeriesToExistingChart(PPT.Chart chart, ChartSeries series)
        {
            var workbook = (EXCEL.Workbook)chart.ChartData.Workbook;
            workbook.Windows.Application.Visible = false;
            var cellNumber = 1;

            var dataSheet = (EXCEL.Worksheet)workbook.Worksheets[1];

            var sc = (PPT.SeriesCollection)chart.SeriesCollection();
            var seriesCount = sc.Count;
            var letter = IntToLetters((seriesCount + 2));

            foreach (var value in series.seriesData)
            {
                var cellPosition = letter + cellNumber.ToString();
                dataSheet.Cells.Range[cellPosition].Value2 = value;
                cellNumber++;
                chart.Refresh();
            }

            // we have to populate the sheet with new values, now we need to create a series for it!
            var columnsRange = dataSheet.UsedRange.Columns;
            var rowsRange = dataSheet.UsedRange.Rows;

            var columnCount = columnsRange.Columns.Count;
            var rowCount = rowsRange.Rows.Count;
            var lastColumnLetter = IntToLetters(columnCount);

            var newSeries = sc.NewSeries();
            newSeries.Name = series.name;
            newSeries.XValues = "'Sheet1'!$A$1:$A$" + rowCount;
            newSeries.Values = "'Sheet1'!$" + lastColumnLetter + "$1:$" + lastColumnLetter + "$" + rowCount;
            newSeries.ChartType = series.seriesType;

            chart.Refresh();
        }

        private static string IntToLetters(int value)
        {
            var result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }
    }
}
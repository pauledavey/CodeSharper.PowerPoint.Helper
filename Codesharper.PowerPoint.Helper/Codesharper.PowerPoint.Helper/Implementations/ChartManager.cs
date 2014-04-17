
namespace Codesharper.PowerPoint.Helper.Implementations
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.Linq;

    using Microsoft.Office.Core;

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using EXCEL = Microsoft.Office.Interop.Excel;

    using Codesharper.PowerPoint.Helper.Contracts;

    public class ChartManager : IChartManager
    {
        //public PPT.Chart CreateChart(PPT.Slide slide, ChartConfiguration chartConfiguration)
        //{
        //    return slide.Shapes.AddChart(
        //            chartConfiguration.chartType,
        //            chartConfiguration.xLocation,
        //            chartConfiguration.yLocation,
        //            chartConfiguration.width,
        //            chartConfiguration.height).Chart;
        //}

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


        public void CreateChart(PPT.Slide slide, string[] xAxisPoints, List<string[]> datasets )
        {
            slide.Layout = PPT.PpSlideLayout.ppLayoutBlank;
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
            for (int i = 1; i != (xAxisPoints.Count() + 1); i++)
            {
                dataSheet.Cells.Range["A" + i].Value2 = xAxisPoints[(i - 1)];
                chart.Refresh();
            }

            var intLetter = 1;
            var cellNumber = 1;

            for (int j = 0; j < datasets.Count; j++)
            {
                var letter = IntToLetters((intLetter + 1));

                // each one of these is a dataset.
                foreach (var value in datasets[j])
                {
                    var cellPosition = letter + cellNumber.ToString();
                    dataSheet.Cells.Range[cellPosition].Value2 = value;
                    cellNumber++;
                    chart.Refresh();
                }

                // we have populate the sheet with new values, now we need to create a series for it!
                EXCEL.Range columnsRange = dataSheet.UsedRange.Columns;
                EXCEL.Range rowsRange = dataSheet.UsedRange.Rows;

                var columnCount = columnsRange.Columns.Count;
                var rowCount = rowsRange.Rows.Count;
                var lastColumnLetter = IntToLetters(columnCount);

                var newSeries = sc.NewSeries();
                newSeries.Name = "Series" + j;
                newSeries.XValues = "'Sheet1'!$A$1:$A$" + rowCount;
                newSeries.Values = "'Sheet1'!$" + lastColumnLetter + "$1:$" + lastColumnLetter + "$" + rowCount;

                if (j == 2)
                {
                    newSeries.ChartType = XlChartType.xlArea;
                }
                else
                {
                    newSeries.ChartType = XlChartType.xlLine;
                }
                

                intLetter++;
                cellNumber = 1;
                chart.Refresh();
            }

            chart.HasTitle = true;
            chart.ChartTitle.Font.Italic = true;
            chart.ChartTitle.Text = "My First Chart!";
            chart.ChartTitle.Font.Size = 12;
            chart.ChartTitle.Font.Color = Color.Black.ToArgb();
            chart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            chart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

            chart.HasLegend = true;
            chart.Legend.Font.Italic = true;
            chart.Legend.Font.Size = 10;

            chart.Refresh();
        }

        public void AddChartTitle(PPT.Shape chart, string titleText)
        {
            chart.Title = titleText;
        }

    }
}




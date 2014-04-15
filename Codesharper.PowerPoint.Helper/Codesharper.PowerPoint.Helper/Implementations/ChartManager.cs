
namespace Codesharper.PowerPoint.Helper.Implementations
{
    using System;
    using System.Drawing;

    using Microsoft.Office.Core;

    using OFFICE = Codesharper.PowerPoint.Helper.Contracts;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using EXCEL = Microsoft.Office.Interop.Excel;

    using Codesharper.PowerPoint.Helper.Contracts;

    public class ChartManager : IChartManager
    {
        public void CreateChart(PPT.Slide slide)
        {
            var chart = slide.Shapes.AddChart(XlChartType.xlLine, 10f, 10f, 900f, 400f).Chart;

            //EXCEL.Workbook workBook = chart.ChartData.Workbook;
            //EXCEL.Worksheet workSheet = workBook.Worksheets.Add();

            var workbook = (EXCEL.Workbook)chart.ChartData.Workbook;
            workbook.Windows.Application.Visible = false;

            var dataSheet = (EXCEL.Worksheet)workbook.Worksheets[1];
            var tRange = dataSheet.Cells.Range["A1", "B5"];
            var tbl1 = dataSheet.ListObjects["Table1"];
            tbl1.Resize(tRange);

            dataSheet.Cells.Range["A2"].FormulaR1C1 = "Bikes";
            dataSheet.Cells.Range["A3"].FormulaR1C1 = "Accessories";
            dataSheet.Cells.Range["A4"].FormulaR1C1 = "Repairs";
            dataSheet.Cells.Range["A5"].FormulaR1C1 = "Clothing";
            dataSheet.Cells.Range["B2"].FormulaR1C1 = "1000";
            dataSheet.Cells.Range["B3"].FormulaR1C1 = "2500";
            dataSheet.Cells.Range["B4"].FormulaR1C1 = "4000";
            dataSheet.Cells.Range["B5"].FormulaR1C1 = "3000";


            PPT.ChartTitle chartTitle = chart.ChartTitle;
            chartTitle.Caption = "My Chart!";
            chartTitle.Position = PPT.XlChartElementPosition.xlChartElementPositionAutomatic;
            chartTitle.Font.Bold = true;
            chartTitle.Font.Italic = true;
            chartTitle.Font.Underline = true;
            chartTitle.Font.Size = 24;
            chartTitle.Font.Color = Color.DeepPink.ToArgb();

            chart.HasLegend = true;

            chart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionBottom;

            chart.ApplyDataLabels(PPT.XlDataLabelsType.xlDataLabelsShowBubbleSizes);

            var mySeries = (PPT.Series)chart.SeriesCollection(1);
            mySeries.Name = "This is a legend!";
            mySeries.MarkerStyle = PPT.XlMarkerStyle.xlMarkerStyleDiamond;
            mySeries.MarkerSize = 10;
            mySeries.Smooth = true;


        }





        public void AddChartTitle(PPT.Shape chart, string titleText)
        {
            chart.Title = titleText;
        }

    }
}

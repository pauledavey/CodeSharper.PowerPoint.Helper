
namespace Codesharper.PowerPoint.Helper.Implementations
{
    using System;
    using System.Diagnostics;
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
            slide.Layout = PPT.PpSlideLayout.ppLayoutBlank;

            var chart = slide.Shapes.AddChart(XlChartType.xlLine, 10f, 10f, 900f, 400f).Chart;
            
            var workbook = (EXCEL.Workbook)chart.ChartData.Workbook;
            workbook.Windows.Application.Visible = true;

            var dataSheet = (EXCEL.Worksheet)workbook.Worksheets[1];
            dataSheet.Cells.ClearContents();
            

            dataSheet.Cells.Range["A1"].Value2 = "Bananas";
            dataSheet.Cells.Range["A2"].Value2 = "Apples";
            dataSheet.Cells.Range["A3"].Value2 = "Pears";
            dataSheet.Cells.Range["A4"].Value2 = "Oranges";
            dataSheet.Cells.Range["B1"].Value2 = "1000";
            dataSheet.Cells.Range["B2"].Value2 = "2500";
            dataSheet.Cells.Range["B3"].Value2 = "4000";
            dataSheet.Cells.Range["B4"].Value2 = "3000";

            var sc = (PPT.SeriesCollection)chart.SeriesCollection();

            do
            {
                var seriesToDelete = sc.Item(1);
                seriesToDelete.Delete();
            }
            while (sc.Count != 0);

            var series1 = sc.NewSeries();
            series1.Name = "Pauls Series";
            series1.XValues = "'Sheet1'!$A$1:$A$2";
            series1.Values = "'Sheet1'!$B$1:$B$2";
            series1.ChartType = XlChartType.xlLine;
           
            var series2 = sc.NewSeries();
            series2.Name = "Seans Series";
            series2.XValues = "'Sheet1'!$A$1:$A$2";
            series2.Values = "'Sheet1'!$B$3:$B$4";
            series2.ChartType = XlChartType.xlLine; 
            
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




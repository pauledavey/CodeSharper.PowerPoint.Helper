namespace Codesharper.PowerPoint.Helper.Specs.PowerPoint.Helper.Specs
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;

    using Codesharper.PowerPoint.Helper.Contracts;
    using Codesharper.PowerPoint.Helper.Implementations;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    using NUnit.Framework;

    using Should;

    using SpecsFor;

    public class when_asked_to_add_a_legend_to_a_chart : SpecsFor<ChartManager>
    {
        private Application powerpointHandle;

        private Presentation presentationHandle;

        private Slide slideHandle;

        private ChartManager chartManager;


        private SlideManager slideManager;

        private PresentationManager presentationManager;

        private Chart returnedChart;

       
        protected override void Given()
        {
            this.chartManager = new ChartManager();
            this.powerpointHandle = new Application();
            this.presentationManager = new PresentationManager();
            this.presentationHandle = this.presentationManager.CreatePowerPointPresentation(this.powerpointHandle, true);
            this.slideManager = new SlideManager();
            this.slideHandle = this.slideManager.AddSlideToEnd(this.presentationHandle);

            var datasetList = new List<ChartSeries>();
            var chartSeries = new ChartSeries()
            {
                name = "Test Series",
                seriesData = new string[] { "10", "20" },
                seriesType = XlChartType.xlLine
            };
            datasetList.Add(chartSeries);

            returnedChart = this.chartManager.CreateChart(XlChartType.xlLine, this.slideHandle, new string[] { "A", "B" }, datasetList);
        }

        protected override void When()
        {
            ChartLegend chartLegend = new ChartLegend()
                                          {
                                                  bold = true,
                                                  italic = true,
                                                  fontSize = 14,
                                                  underline = true
                                          };

            this.SUT.AddChartLegend(returnedChart, chartLegend);
        }

        [Test]
        public void then_we_should_get_a_chart_with_a_legend()
        {
            this.returnedChart.HasLegend.ShouldBeTrue();
        }

        protected override void AfterEachTest()
        {
            base.AfterEachTest();
            this.powerpointHandle.Quit();
        }
    }
}

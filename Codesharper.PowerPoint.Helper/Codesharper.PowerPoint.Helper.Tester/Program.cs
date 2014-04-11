namespace Codesharper.PowerPoint.Helper.Tester
{
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using Codesharper.PowerPoint.Helper.Implementations;
    using PPTHelper = Codesharper.PowerPoint.Helper.Implementations.Presentation;
    using PPTApplication = Codesharper.PowerPoint.Helper.Implementations.PowerPointApplication;
    using PPTSlideHelper = Codesharper.PowerPoint.Helper.Implementations.SlideManager;


    class Program
    {
        static void Main(string[] args)
        {
            // Initial object setups
            var helper = new PPTHelper();
            var pptAppHelper = new PPTApplication();
            var pptApp = pptAppHelper.CreatePowerPointApplication();
            var pptSlideHelper = new PPTSlideHelper();
            
            var shapeHelper = new Shapes();

            //create a new PPT application instance
            var presentation = helper.CreatePowerPointPresentation(pptApp, false);

            // add slide to the end of the presentation
            var slideAtEnd = pptSlideHelper.AddSlideAtEndOfPresentation(presentation);

            // insert slide in to presentation
            pptSlideHelper.InsertSlideIntoPresentation(presentation, (pptSlideHelper.GetSlideCountInPresentation(presentation) + 1));
            
            // grab the first slide in the presentation
            var mySlide = presentation.Slides[1];

            //// AddTextBox to the the first slide and set some text
            //var textbox = shapeHelper.AddTextBoxToSlide(mySlide, OFFICE.MsoTextOrientation.msoTextOrientationHorizontal, 100f, 100f, 250f, 250f);
            //shapeHelper.SetTextBoxText(textbox, "Goat Love Rules Man! YEAH!");

            //var textbox2 = shapeHelper.AddTextBoxToSlide(mySlide, OFFICE.MsoTextOrientation.msoTextOrientationDownward, 400f, 300f, 350f, 350f);
            //shapeHelper.SetTextBoxText(textbox2, "Going Down!!!");

            //// var chart1 = shapeHelper.AddChartToSlide(mySlide, OFFICE.XlChartType.xlLine, 50f, 50f, 50f, 50f);
            //var objChart = (Graph.Chart)mySlide.Shapes.AddOLEObject(150, 150, 480, 320, "MSGraph.Chart.8", "", OFFICE.MsoTriState.msoFalse, "", 0, "", OFFICE.MsoTriState.msoFalse).OLEFormat.Object;
            //objChart.ChartType = Graph.XlChartType.xl3DPie;
            //objChart.Legend.Position = Graph.XlLegendPosition.xlLegendPositionBottom;
            //objChart.HasTitle = true;
            //objChart.ChartTitle.Text = "Goats I have loved and known!";

            //try
            //{
            //    mySlide.Layout = PPT.PpSlideLayout.ppLayoutChartAndText;
            //    var objChart = mySlide.Shapes.AddChart(OFFICE.XlChartType.xlLine, 100, 100, 100, 100);
            //    objChart.Title = "Goats I have loved, lost and known";
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception("Your Goat has crashed. Stop bleeting and stand it up..");
            //}


            var objTable = mySlide.Shapes.AddTable(10, 3, 100f, 100f, 300f, 300f);



            // save and close powerpoint
            helper.SavePresentationAs(presentation, @"c:\example.ppt", PPT.PpSaveAsFileType.ppSaveAsPresentation, true);
            helper.ClosePresentation(presentation);
        }
    }
}

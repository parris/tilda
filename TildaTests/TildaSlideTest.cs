using Tilda.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using TildaTests.Mocks;
using Office = Microsoft.Office.Core;

namespace TildaTests
{
    
    
    /// <summary>
    ///This is a test class for TildaSlideTest and is intended
    ///to contain all TildaSlideTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaSlideTest {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext {
            get {
                return testContextInstance;
            }
            set {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for TildaSlide Constructor
        ///</summary>
        [TestMethod()]
        public void TildaSlideConstructorTest() {
            Slide slide = new TildaTests.Mocks.MockSlide();
            TildaSlide target = new TildaSlide(slide);
            Assert.AreEqual(slide, target.slide);
        }

        /// <summary>
        /// Tests that function is created
        ///</summary>
        [TestMethod()]
        public void slideJSExportCode() {
            if(Directory.Exists(Settings.outputPath))
                Directory.Delete(Settings.outputPath, true);
            Directory.CreateDirectory(Settings.outputPath);
            Directory.CreateDirectory(Settings.outputMediaFullPath);

            Slide slide = new TildaTests.Mocks.MockSlide();
            TildaSlide target = new TildaSlide(slide);
            String actual = target.exportSlide();

            Assert.AreEqual(true, actual.Contains("function(){"));
            Assert.AreEqual(true, actual[actual.Length-1]=='}');

            Directory.Delete(Settings.outputPath, true);
        }

        /// <summary>
        /// Tests that a background image is exported to the right location
        ///</summary>
        [TestMethod()]
        public void backgroundImageExported() {
            if(Directory.Exists(Settings.outputPath))
                Directory.Delete(Settings.outputPath, true);
            Directory.CreateDirectory(Settings.outputPath);
            Directory.CreateDirectory(Settings.outputMediaFullPath);

            Slide slide = new TildaTests.Mocks.MockSlide(); 
            TildaSlide target = new TildaSlide(slide); 
            String actual = target.exportSlide();

            bool bgcreated = false;
            foreach(string file in Directory.GetFiles(Settings.outputMediaFullPath)){
                if(file.Contains("-bg.png"))
                    bgcreated = true;
            }

            Assert.AreEqual(true, bgcreated);
            Assert.AreEqual(true, actual.Contains("preso.shapes.push(preso.paper.image('assets/"));
            Assert.AreEqual(true, actual.Contains(",0,0,2160,3800));preso.shapes[(preso.shapes.length-1)].toBack();"));

            Directory.Delete(Settings.outputPath, true);
        }

        /// <summary>
        /// Z-order fixes shape order
        ///</summary>
        [TestMethod()]
        public void shapesAreOrderedBasedOnZOrderAndVisibilityIsNotConsidered() {
            if(Directory.Exists(Settings.outputPath))
                Directory.Delete(Settings.outputPath, true);
            Directory.CreateDirectory(Settings.outputPath);
            Directory.CreateDirectory(Settings.outputMediaFullPath);

            Slide slide = new TildaTests.Mocks.MockSlide();
            TildaSlide target = new TildaSlide(slide);
            Shape tb = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward, 0, 0, 100, 100);
            tb.TextFrame2.TextRange.Text = "Hello2";

            Shape tb2 = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward, 0, 0, 100, 100);
            tb2.TextFrame2.TextRange.Text = "Hello1";
            tb2.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

            Shape tb3 = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward, 0, 0, 100, 100);
            tb3.TextFrame2.TextRange.Text = "Hello3";
            tb3.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            tb3.Visible = Office.MsoTriState.msoFalse;

            String actual = target.exportSlide();

            int first = actual.IndexOf("Hello1");
            int second = actual.IndexOf("Hello2");
            int third = actual.IndexOf("Hello3");

            Assert.AreEqual(true, first < second);
            Assert.AreEqual(true, third < second);

            Directory.Delete(Settings.outputPath, true);
        }

        //TODO: Shapes exported based on type
        //TODO: Animation considerations
    }
}

using WLB_Builder.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;
using WBLTests.Mocks;

namespace WBLTests
{
    /// <summary>
    ///This is a test class for WBLSlideTest and is intended
    ///to contain all WBLSlideTest Unit Tests
    ///</summary>
    [TestClass()]
    public class WBLSlideTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
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
        ///A test for WBLSlide Constructor
        ///</summary>
        [TestMethod()]
        public void WBLSlideConstructorTest()
        {
            Slide slide = new Mocks.MockSlide();
            WBLSlide target = new WBLSlide(slide);
            Assert.AreEqual(slide, target.slide);
        }

        /// <summary>
        ///A test for exportSlide
        ///</summary>
        [TestMethod()]
        public void exportSlideTest()
        {
            Slide slide = new Mocks.MockSlide();
            WBLSlide target = new WBLSlide(slide);

            String html = target.exportSlide("/");
            //Assert.AreEqual("<div id=\"ppt-slide\"></div>", html);
            Assert.AreEqual("", html);

            Shape shape1 = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward, 10, 2, 100, 100);
            shape1.Rotation = 90;
            shape1.TextEffect.FontName = "Arial";
            shape1.TextEffect.FontSize = 16;
            shape1.TextEffect.Text = "Title";
            shape1.Fill.ForeColor.RGB = 0;

            //compare
            String actual = target.exportSlide("/");
            /*String font = "font-style:Arial;font-size:16px;color:#000000;";
            String deg = "-moz-transform:rotate(90deg);-webkit-transform:rotate(90deg);-o-transform:rotate(90deg);-ms-transform:rotate(90deg);"
                      + "filter:progid:DXImageTransform.Micrsoft.BasicImage(rotation=1);";
            String pos = "top:" + shape1.Top + "px;left:" + shape1.Left + "px;width:" + shape1.Width + "px;height:" + shape1.Height + "px;";
            String expected = "<div id=\"ppt-slide\"><div class=\"ppt-textbox\" style=\"" + font + pos + deg + "\">Title<div></div>";*/

            string expected = "var textbox-5 = paper.text(2,10,'').attr({'font-style':'Arial','font-size':'16','color':'#000000','transformation':'r90'});";

            Assert.AreEqual(expected, actual);
        }
    }
}

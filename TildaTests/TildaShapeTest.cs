using Tilda.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;
using WBLTests.Mocks;
using System.Drawing;

namespace Tilda.Tests
{
    
    
    /// <summary>
    ///This is a test class for TildaShapeTest and is intended
    ///to contain all TildaShapeTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaShapeTest
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
        ///A test for TildaShape Constructor
        ///</summary>
        [TestMethod()]
        public void TildaShapeConstructorTest()
        {
            Shape shape = new MockShape(); 
            TildaShape target = new TildaShape(shape);
            Assert.AreEqual(shape, target.shape);
        }

        /// <summary>
        ///A test for cssFont
        ///</summary>
        [TestMethod()]
        public void cssFontTest()
        {
            //create Shape
            Shape shape = new MockShape();
            //set attributes
            shape.TextEffect.FontName = "Arial";
            shape.TextEffect.FontSize = 16;
            shape.Fill.ForeColor.RGB = 0;

            //compare
            //String expected = "font-style:Arial;font-size:16px;color:#000000;";
            String expected = "'font-style':'Arial','font-size':'16','color':'#000000'";
            TildaShape target = new TildaShape(shape);
            String actual = target.fontStyle();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for cssRotation
        ///</summary>
        [TestMethod()]
        public void cssRotationTest()
        {
            //create Shape
            Shape shape = new MockShape();
            shape.Rotation = 90;
            //set attributes

            //compare
            /*var expected = "-moz-transform:rotate(90deg);-webkit-transform:rotate(90deg);-o-transform:rotate(90deg);-ms-transform:rotate(90deg);"
                      + "filter:progid:DXImageTransform.Micrsoft.BasicImage(rotation=1);";*/
            String expected = "'transformation':'r90'";
            TildaShape target = new TildaShape(shape);
            String actual = target.transformation();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for cssSizePosition
        ///</summary>
        [TestMethod()]
        public void cssSizePositionTest()
        {
            //create Shape
            Shape shape = new MockShape();
            shape.Width = 100;
            shape.Height = 100;
            shape.Top = 2;
            shape.Left = 10;
            //set attributes

            //compare
            var expected = "top:" + shape.Top + "px;left:" + shape.Left + "px;width:" + shape.Width + "px;height:" + shape.Height + "px;";
            TildaShape target = new TildaShape(shape);
            String actual = target.cssSizePosition();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for extractTextBox
        ///</summary>
        [TestMethod()]
        public void extractTextBoxTest()
        {
            //create Shape
            Shape shape = new MockShape();
            //set attributes
            shape.Width = 100;
            shape.Height = 100;
            shape.Top = 2;
            shape.Left = 10;
            shape.Rotation = 90;
            shape.TextEffect.FontName = "Arial";
            shape.TextEffect.FontSize = 16;
            shape.TextEffect.Text = "Title";
            shape.Fill.ForeColor.RGB = 0;

            //compare
            /*String font = "font-style:Arial;font-size:16px;color:#000000;";
            String deg = "-moz-transform:rotate(90deg);-webkit-transform:rotate(90deg);-o-transform:rotate(90deg);-ms-transform:rotate(90deg);"
                      + "filter:progid:DXImageTransform.Micrsoft.BasicImage(rotation=1);";
            String pos = "top:" + shape.Top + "px;left:" + shape.Left + "px;width:" + shape.Width + "px;height:" + shape.Height + "px;";
            String expected = "<div class=\"ppt-textbox\" style=\"" + font + pos + deg + "\">Title<div>";*/
            String expected = "var testing = paper.text(2,10,'').attr({'font-style':'Arial','font-size':'16','color':'#000000','transformation':'r90'});";
            TildaShape target = new TildaShape(shape);
            String actual = target.readText();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for rgbToHex
        ///</summary>
        [TestMethod()]
        public void rgbToHexTest()
        {
            TildaShape s = new TildaShape(new MockShape());
            Assert.AreEqual("#ffffff", s.rgbToHex(16777215));
            Assert.AreEqual("#7b72de", s.rgbToHex(8090334));
            Assert.AreEqual("#e772de", s.rgbToHex(15168222));
        }

        /// <summary>
        ///A test for extractMP3
        ///</summary>
        [TestMethod()]
        public void extractMP3Test()
        {
            Shape shape = null; // TODO: Initialize to an appropriate value
            TildaShape target = new TildaShape(shape); // TODO: Initialize to an appropriate value
            string directory = string.Empty; // TODO: Initialize to an appropriate value
            //bool expected = false; // TODO: Initialize to an appropriate value
            //bool actual;
            //actual = target.extractMP3(directory);
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }
        
        /// <summary>
        ///A test for name
        ///</summary>
        [TestMethod()]
        public void nameTest()
        {
            TildaShape shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoAutoShape));
            Assert.AreEqual(true,shape.name().Contains("autoshape"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoCallout));
            Assert.AreEqual(true, shape.name().Contains("callout"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoCanvas));
            Assert.AreEqual(true, shape.name().Contains("canvas"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoChart));
            Assert.AreEqual(true, shape.name().Contains("chart"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoComment));
            Assert.AreEqual(true, shape.name().Contains("comment"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoDiagram));
            Assert.AreEqual(true, shape.name().Contains("diagram"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject));
            Assert.AreEqual(true, shape.name().Contains("embeddedoleobj"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoFormControl));
            Assert.AreEqual(true, shape.name().Contains("formcontrol"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoFreeform));
            Assert.AreEqual(true, shape.name().Contains("freeform"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoGroup));
            Assert.AreEqual(true, shape.name().Contains("group"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoInk));
            Assert.AreEqual(true, shape.name().Contains("ink"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoInkComment));
            Assert.AreEqual(true, shape.name().Contains("inkcomment"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoLine));
            Assert.AreEqual(true, shape.name().Contains("line"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject));
            Assert.AreEqual(true, shape.name().Contains("linkedoleobj"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoLinkedPicture));
            Assert.AreEqual(true, shape.name().Contains("linkedpicture"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoMedia));
            Assert.AreEqual(true, shape.name().Contains("media"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoOLEControlObject));
            Assert.AreEqual(true, shape.name().Contains("olecontrolobject"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoPicture));
            Assert.AreEqual(true, shape.name().Contains("picture"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoPlaceholder));
            Assert.AreEqual(true, shape.name().Contains("placeholder"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoScriptAnchor));
            Assert.AreEqual(true, shape.name().Contains("scriptanchor"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoShapeTypeMixed));
            Assert.AreEqual(true, shape.name().Contains("shapetypemixed"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoSlicer));
            Assert.AreEqual(true, shape.name().Contains("slicer"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoSmartArt));
            Assert.AreEqual(true, shape.name().Contains("smartart"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoTable));
            Assert.AreEqual(true, shape.name().Contains("table"));
            shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoTextBox));
            Assert.AreEqual(true, shape.name().Contains("textbox"));
        }
    }
}

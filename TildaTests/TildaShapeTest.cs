using Tilda.Models;
using TildaTests;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;
using System.Drawing;

namespace TildaTests
{
    
    
    /// <summary>
    ///This is a test class for TildaShapeTest and is intended
    ///to contain all TildaShapeTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaShapeTest {


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
        ///A test for TildaShape Constructor
        ///</summary>
        [TestMethod()]
        public void TildaShapeConstructorTest() {
            Shape shape = new MockShape();
            TildaShape target = new TildaShape(shape,0);
            Assert.AreEqual(shape, target.shape);
            Assert.AreEqual(Settings.Scaler(), target.scaler);
            Assert.AreEqual(0, target.id);
        }

        /// <summary>
        ///A test for cssRotation
        ///</summary>
        [TestMethod()]
        public void RotationTest() {
            //create Shape
            Shape shape = new MockShape();
            shape.Rotation = 90;
            //set attributes

            //compare
            String expected = "'transformation':'r90'";
            TildaShape target = new TildaShape(shape);
            String actual = target.transformation();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for rgbToHex
        ///</summary>
        [TestMethod()]
        public void rgbToHexTest() {
            TildaShape s = new TildaShape(new MockShape());
            Assert.AreEqual("#ffffff", s.rgbToHex(16777215));
            Assert.AreEqual("#de727b", s.rgbToHex(8090334));
            Assert.AreEqual("#de72e7", s.rgbToHex(15168222));
        }

        /// <summary>
        ///A test for name
        ///</summary>
        [TestMethod()]
        public void nameTest() {
            TildaShape shape = new TildaShape(new MockShape(Microsoft.Office.Core.MsoShapeType.msoAutoShape));
            Assert.AreEqual(true, shape.name().Contains("autoshape"));
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

        /// <summary>
        ///A test for findX
        ///</summary>
        [TestMethod()]
        public void findXTest() {
            Shape shape = new MockShape();
            shape.Left = 50.3f;
            int id = 0; 
            TildaShape target = new TildaShape(shape, id);
            double expected = shape.Left * Settings.Scaler(); 
            double actual;
            actual = target.findX();
            double difference = expected - actual;
            Assert.IsTrue(difference <= .00001);
        }

        /// <summary>
        ///A test for findY
        ///</summary>
        [TestMethod()]
        public void findYTest() {
            Shape shape = new MockShape();
            shape.Top = 50.3f;
            int id = 0;
            TildaShape target = new TildaShape(shape, id);
            double expected = shape.Top * Settings.Scaler();
            double actual;
            actual = target.findY();
            double difference = expected - actual;
            Assert.IsTrue(difference <= .00001);
        }

        /// <summary>
        ///A test for position
        ///</summary>
        [TestMethod()]
        public void positionTest() {
            Shape shape = new MockShape();
            shape.Left = 30f;
            shape.Top = 55f;
            int id = 5; 
            TildaShape target = new TildaShape(shape, id);
            string expected = (double)(shape.Left * Settings.Scaler()) + "," + (double)(shape.Top * Settings.Scaler()); 
            string actual;
            actual = target.position();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for toRaphJS
        ///</summary>
        [TestMethod()]
        public void toRaphJSTest() {
            Shape shape = new MockShape();
            int id = 0; 
            TildaShape target = new TildaShape(shape, id);
            TildaAnimation[] animationMap = null; 
            TildaSlide slide = null;
            string expected = string.Empty;
            string actual;
            actual = target.toRaphJS(animationMap, slide);
            Assert.AreEqual(expected, actual);
        }
    }
}

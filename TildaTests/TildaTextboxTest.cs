using Tilda.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests
{
    
    
    /// <summary>
    ///This is a test class for TildaTextboxTest and is intended
    ///to contain all TildaTextboxTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaTextboxTest {


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
        ///A test for TildaTextbox Constructor
        ///</summary>
        [TestMethod()]
        public void TildaTextboxConstructorTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id);
            //Assert.Inconclusive("TODO: Implement code to verify target");
        }

        /// <summary>
        ///Overriden findX function for Textboxes
        ///</summary>
        [TestMethod()]
        public void findXTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //double expected = 0F; // TODO: Initialize to an appropriate value
            double actual;
            //actual = target.findX();
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for findY
        ///</summary>
        [TestMethod()]
        public void findYTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //double expected = 0F; // TODO: Initialize to an appropriate value
            //double actual;
            //actual = target.findY();
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for fontPosition
        ///</summary>
        [TestMethod()]
        public void fontPositionTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //float addx = 0F; // TODO: Initialize to an appropriate value
            //float addy = 0F; // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = target.fontPosition(addx, addy);
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for fontStyle
        ///</summary>
        [TestMethod()]
        public void fontStyleTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = target.fontStyle();
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for position
        ///</summary>
        [TestMethod()]
        public void positionTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //float xOffset = 0F; // TODO: Initialize to an appropriate value
            //float yOffset = 0F; // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = target.position(xOffset, yOffset);
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for getParagraphs
        ///</summary>
        [TestMethod()]
        public void tildifyTextTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = target.getParagraphs();
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for toRaphJS
        ///</summary>
        [TestMethod()]
        public void toRaphJSTest() {
            Shape shape = null; // TODO: Initialize to an appropriate value
            int id = 0; // TODO: Initialize to an appropriate value
            //TildaTextbox target = new TildaTextbox(shape, id); // TODO: Initialize to an appropriate value
            //TildaAnimation[] animationMap = null; // TODO: Initialize to an appropriate value
            //TildaSlide slide = null; // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = target.toRaphJS(animationMap, slide);
            //Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}

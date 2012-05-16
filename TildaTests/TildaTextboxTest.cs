using Tilda.Models;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;

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
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            TildaTextbox target = new TildaTextbox(shape, id);
            Assert.AreEqual(shape,target.shape);
            Assert.AreEqual(id, target.id);
        }

        /// <summary>
        ///Overriden findX function for Textboxes
        ///</summary>
        [TestMethod()]
        public void findXTest() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            MockTextRange2 tr = (MockTextRange2)shape.TextFrame2.TextRange;
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            TildaTextbox target = new TildaTextbox(shape, id);

            //left aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            Assert.AreEqual(Math.Round((shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler()), target.findX());

            //right aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight;
            Assert.AreEqual(Math.Round((shape.Left + shape.Width - shape.TextFrame2.MarginRight) * Settings.Scaler()), target.findX());

            //centered
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter;
            Assert.AreEqual(Math.Round((shape.Width/2 + shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler()),target.findX());

            //dont care otherwise
        }

        /// <summary>
        ///A test for findY
        ///</summary>
        [TestMethod()]
        public void findYTest() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            MockTextRange2 tr = (MockTextRange2)shape.TextFrame2.TextRange;
            shape.Top = 5f;
            shape.TextFrame2.MarginTop = 3.2f;
            shape.TextFrame2.MarginBottom = 4.2f;
            shape.Height = 100f;
            TildaTextbox target = new TildaTextbox(shape, id);

            //bottom aligned
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom;
            Assert.AreEqual((shape.Top + shape.Height - shape.TextFrame2.MarginBottom) * Settings.Scaler(), target.findY(),.001);

            //baseline aligned
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottomBaseLine;
            Assert.AreEqual((shape.Top + shape.Height) * Settings.Scaler(), target.findY(), .001);

            //top aligned
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
            Assert.AreEqual((shape.Top + shape.TextFrame2.MarginTop) * Settings.Scaler(), target.findY(), .001);

            //top baseline aligned
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTopBaseline;
            Assert.AreEqual((shape.Top) * Settings.Scaler(), target.findY(), .001);

            //middle
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            Assert.AreEqual((shape.Height / 2 + shape.TextFrame2.MarginTop + shape.Top) * Settings.Scaler(), target.findY(), .001);
            

            //otherwise don't care!
        }

        /// <summary>
        ///A test for fontPosition
        ///</summary>
        [TestMethod()]
        public void fontPositionTest() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            MockTextRange2 tr = (MockTextRange2)shape.TextFrame2.TextRange;
            List<MockTextRange2> pgs = new List<MockTextRange2>();
            pgs.Add(new MockTextRange2("Parapgrah1"));
            pgs.Add(new MockTextRange2("Paragraph2"));
            tr.set_Paragraphs(pgs);
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            TildaTextbox target = new TildaTextbox(shape, id);
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            Assert.AreEqual(target.findX(), (shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler());
        }

        /// <summary>
        ///A test for fontStyle
        ///</summary>
        [TestMethod()]
        public void fontStyleTest() {
            PowerPoint.Shape shape = null; // TODO: Initialize to an appropriate value
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
            PowerPoint.Shape shape = null; // TODO: Initialize to an appropriate value
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
            PowerPoint.Shape shape = null; // TODO: Initialize to an appropriate value
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
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            MockTextRange2 tr = (MockTextRange2)shape.TextFrame2.TextRange;
            List<MockTextRange2> pgs = new List<MockTextRange2>();
            pgs.Add(new MockTextRange2("Parapgrah1"));
            pgs.Add(new MockTextRange2("Paragraph2"));
            tr.set_Paragraphs(pgs);
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            TildaTextbox target = new TildaTextbox(shape, id);
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            Assert.AreEqual(target.findX(), (shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler());
        }
    }
}

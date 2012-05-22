using Tilda.Models;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
            Assert.AreEqual((shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler(), target.findX(), .001);

            //right aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight;
            Assert.AreEqual((shape.Left + shape.Width - shape.TextFrame2.MarginRight) * Settings.Scaler(), target.findX(), .001);

            //centered
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter;
            Assert.AreEqual((shape.Width / 2 + shape.Left + shape.TextFrame2.MarginLeft) * Settings.Scaler(), target.findX(), .001);

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
            TildaTextbox target = new TildaTextbox(shape, id);

            //left aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            Assert.AreEqual("'text-anchor': 'start'", target.fontPosition());

            //center aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter;
            Assert.AreEqual("'text-anchor': 'middle'", target.fontPosition());

            //right aligned
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight;
            Assert.AreEqual("'text-anchor': 'end'", target.fontPosition());

            //otherwise don't care
        }

        /// <summary>
        ///A test for fontStyle
        ///</summary>
        [TestMethod()]
        public void fontStyleTest() {
            int redRGB = 16711680;
            string redHex = "#ff0000";

            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            MockTextRange2 tr = (MockTextRange2)shape.TextFrame2.TextRange;
            tr.Font.Name = "Verdana";
            tr.Font.Size = 12f;
            tr.Font.Fill.ForeColor.RGB = redRGB;
            TildaTextbox target = new TildaTextbox(shape, id);

            String expected = "'font-size':'" + Settings.Scaler() * tr.Font.Size + "','fill':'" + redHex + "'";
            //No bold or italic
            Assert.AreEqual(expected + ",'font-family':'" + tr.Font.Name + "'", target.fontStyle());

            //Italic only
            tr.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;
            Assert.AreEqual(expected + ",'font-family':'" + tr.Font.Name + " italic'", target.fontStyle());
            tr.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue;
            Assert.AreEqual(expected + ",'font-family':'" + tr.Font.Name + " italic'", target.fontStyle());

            //Bold only
            tr.Font.Italic = Microsoft.Office.Core.MsoTriState.msoFalse;
            tr.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            Assert.AreEqual(expected + ",'font-weight':'bold','font-family':'" + tr.Font.Name + "'", target.fontStyle());
            tr.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
            Assert.AreEqual(expected + ",'font-weight':'bold','font-family':'" + tr.Font.Name + "'", target.fontStyle());

            //Bold and Italic
            tr.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;
            Assert.AreEqual(expected + ",'font-weight':'bold','font-family':'" + tr.Font.Name + " italic'", target.fontStyle());
        }

        /// <summary>
        /// Tests to see if 1 paragraph can be generated, 1 line
        ///</summary>
        [TestMethod()]
        public void canGenerateOneParagraph() {
            TildaTextbox tb = oneSentenceFixture();
            int fontsize = (int)(tb.shape.TextFrame2.TextRange.Font.Size * Settings.Scaler());
            //no need to check these functions, but we will be adding values to them
            double x = tb.findX();
            double y = tb.findY();

            // offsetting text
            y += (tb.shape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore) * Settings.Scaler();

            String expected = "idsToAnimate = new Array();"+
                "preso.shapes.push(preso.paper.text("+x+","+y+",'Paragraph1').attr("
                + "{'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));";

            String actual = tb.toRaphJS();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// Tests to see if 1 paragraph can be generate, 3 lines
        ///</summary>
        [TestMethod()]
        public void canGenerateOneParagraphWithMultipleLines() {
            TildaTextbox tb = multiLineSentenceFixture();

            int fontsize = (int)(tb.shape.TextFrame2.TextRange.Font.Size * Settings.Scaler());

            double x = tb.findX();
            Microsoft.Office.Core.ParagraphFormat2 pgformat = tb.shape.TextFrame2.TextRange.ParagraphFormat;
            double y1 = tb.findY() + (pgformat.SpaceBefore) * Settings.Scaler();
            double y2 = y1 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();
            double y3 = y2 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();

            String expected = "idsToAnimate = new Array();preso.shapes.push("+
                "preso.paper.text(" + x + "," + y1 + ",'Line1').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y2 + ",'Line2').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y3 + ",'This is Line the 3rd').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));";

            String actual = tb.toRaphJS();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// Tests to see if 2 paragraph can be generate, 3 lines
        ///</summary>
        [TestMethod()]
        public void canGenerateMultipleParagraphsWithMultipleLines() {
            TildaTextbox tb = multiLineAndParagraphFixture();

            int fontsize = (int)(tb.shape.TextFrame2.TextRange.Font.Size * Settings.Scaler());

            double x = tb.findX();
            Microsoft.Office.Core.ParagraphFormat2 pgformat = tb.shape.TextFrame2.TextRange.ParagraphFormat;
            double y1 = tb.findY() + (pgformat.SpaceBefore) * Settings.Scaler();
            double y2 = y1 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();
            double y3 = y2 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();
            double y4 = y3 + fontsize + (pgformat.SpaceAfter + pgformat.SpaceBefore) * Settings.Scaler();
            double y5 = y4 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();
            double y6 = y5 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler();

            String expected = "idsToAnimate = new Array();"+
                "preso.shapes.push(preso.paper.text(" + x + "," + y1 + ",'Paragraph1Line1').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y2 + ",'Line2').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y3 + ",'Line3').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "idsToAnimate = new Array();"+
                "preso.shapes.push(preso.paper.text(" + x + "," + y4 + ",'Paragraph2Line3').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y5 + ",'Line4').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(" + x + "," + y6 + ",'Line5').attr({'font-size':'" + fontsize + "','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));";

            String actual = tb.toRaphJS();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// Tests to see if a few bullets can be created
        ///</summary>
        [TestMethod()]
        public void canGenerateSingleLevelBullets() {
            TildaTextbox tb = singleLevelBulletsFixture();

            int fontsize = (int)(tb.shape.TextFrame2.TextRange.Font.Size * Settings.Scaler());
            double bulletPadding = fontsize / 8;

            double x = tb.findX();
            Microsoft.Office.Core.ParagraphFormat2 pgformat = tb.shape.TextFrame2.TextRange.ParagraphFormat;
            double y1 = tb.findY() + (pgformat.SpaceBefore) * Settings.Scaler() + bulletPadding;
            double y2 = y1 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler() + bulletPadding * 2;
            double y3 = y2 + fontsize + (pgformat.SpaceAfter) * Settings.Scaler() + bulletPadding * 2;

            String expected = "idsToAnimate = new Array();" +
                "preso.shapes.push(preso.paper.rect(62.2000007629395,10.3999996185303,12,12).attr({'stroke':'#ff0000','fill':'#ff0000'}));" +
                "preso.shapes.push(preso.paper.text(38.2000007629395,13.3999996185303,'bullet1').attr({'font-size':'24','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(38.2000007629395,50.7999992370605,'Bullet2').attr({'font-size':'24','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));" +
                "preso.shapes.push(preso.paper.text(38.2000007629395,88.1999988555908,'Bullet3').attr({'font-size':'24','fill':'#ff0000','font-family':'Verdana','transformation':'r0','text-anchor': 'start'}));";
            String actual = tb.toRaphJS();
            //Assert.AreEqual(expected, actual);
        }

        //some fixtures, these should prob end up in their own files. Not quite sure what to do with them in C# yet.
        private TildaTextbox oneSentenceFixture() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            Microsoft.Office.Core.TextRange2 tr = shape.TextFrame2.TextRange; 
            tr.Text = "Paragraph1";
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            tr.Font.Name = "Verdana";
            tr.Font.Size = 12f;
            tr.ParagraphFormat.SpaceBefore = 5.2f;
            int redRGB = 16711680;
            tr.Font.Fill.ForeColor.RGB = redRGB;
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
            return new TildaTextbox(shape, id);
        }

        private TildaTextbox multiLineSentenceFixture() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            Microsoft.Office.Core.TextRange2 tr = shape.TextFrame2.TextRange;
            tr.Text = "Line1~Line2~This is Line the 3rd";
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            tr.Font.Name = "Verdana";
            tr.Font.Size = 12f;
            tr.ParagraphFormat.SpaceBefore = 5.2f;
            tr.ParagraphFormat.SpaceAfter = 5.2f;
            int redRGB = 16711680;
            tr.Font.Fill.ForeColor.RGB = redRGB;
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
            return new TildaTextbox(shape, id);
        }

        private TildaTextbox multiLineAndParagraphFixture() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            Microsoft.Office.Core.TextRange2 tr = shape.TextFrame2.TextRange;
            tr.Text = "Paragraph1Line1~Line2~Line3\rParagraph2Line3~Line4~Line5";
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            tr.Font.Name = "Verdana";
            tr.Font.Size = 12f;
            tr.ParagraphFormat.SpaceBefore = 5.2f;
            tr.ParagraphFormat.SpaceAfter = 5.2f;
            int redRGB = 16711680;
            tr.Font.Fill.ForeColor.RGB = redRGB;
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
            return new TildaTextbox(shape, id);
        }

        private TildaTextbox singleLevelBulletsFixture() {
            PowerPoint.Shape shape = new MockShape();
            int id = 15;
            Microsoft.Office.Core.TextRange2 tr = shape.TextFrame2.TextRange;
            // ` character is a level 1 bullet when the mock renders
            tr.Text = "`Bullet1~`Bullet2~`Bullet3";
            shape.Left = 7f;
            shape.Width = 100f;
            shape.TextFrame2.MarginLeft = 1.1f;
            shape.TextFrame2.MarginRight = 1.2f;
            tr.Font.Name = "Verdana";
            tr.Font.Size = 12f;
            tr.ParagraphFormat.SpaceBefore = 5.2f;
            tr.ParagraphFormat.SpaceAfter = 5.2f;
            int redRGB = 16711680;
            tr.Font.Fill.ForeColor.RGB = redRGB;
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
            return new TildaTextbox(shape, id);
        }
    }
}

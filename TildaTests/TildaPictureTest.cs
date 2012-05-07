using Tilda.Models;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace TildaTests
{
    
    
    /// <summary>
    ///This is a test class for TildaPictureTest and is intended
    ///to contain all TildaPictureTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaPictureTest {


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
        ///A test for TildaPicture Constructor
        ///</summary>
        [TestMethod()]
        public void TildaShapeConstructorTest() {
            Shape shape = new MockShape();
            int id = 0;
            TildaPicture target = new TildaPicture(shape, id);
            Assert.AreEqual(shape, target.shape);
            Assert.AreEqual(Settings.Scaler(), target.scaler);
            Assert.AreEqual(0, target.id);
        }

        /// <summary>
        ///A test for toRaphJS
        ///</summary>
        [TestMethod()]
        public void toRaphJSTest() {
            Shape shape = new MockShape();
            int id = 0; 
            TildaPicture target = new TildaPicture(shape, id); 

            TildaShape[] shapeMap = new TildaShape[2];
            shapeMap[0] = target;
            shapeMap[0].shape.Width = 20f;
            shapeMap[0].shape.Height = 40f;
            shapeMap[0].shape.Top = 5f;
            shapeMap[0].shape.Left = 6f;
            shapeMap[1] = new TildaShape(new MockShape(), 1);
            shapeMap[1].shape.Width = 50f;
            shapeMap[1].shape.Height = 60f;

            List<TildaAnimation> animationMap = new List<TildaAnimation>(); 

            TildaSlide slide = new TildaSlide(new TildaTests.Mocks.MockSlide());
            string expected = @"preso.shapes.push\(preso.paper.image\('assets/[0-9]*-[0-9]*-image.png',"
                + shapeMap[0].position() + ","+(shapeMap[0].shape.Width*Settings.Scaler())+","+shapeMap[0].shape.Height*Settings.Scaler()+@"\)\);";
            string actual;

            //Assert.AreEqual(slide.shapeCount, 0);
            actual = target.toRaphJS();

            Boolean doesEqual = Regex.IsMatch(actual,expected);
            Assert.AreEqual(true, doesEqual);

            //Adding animations
            TildaAnimation anim = new TildaAnimation(new MockEffect(),shapeMap[0]);
            anim.effect.Timing.Duration = 5f;
            anim.effect.Timing.TriggerDelayTime = 15f;
            shapeMap[0].animations.Add(anim);

            expected = @"preso.shapes.push\(preso.paper.image\('assets/[0-9]*-[0-9]*-image.png',"
                + shapeMap[0].position() + "," + (shapeMap[0].shape.Width * Settings.Scaler()) + "," + shapeMap[0].shape.Height * Settings.Scaler() + @"\)\);"
                + @"preso.shapes\[\(preso.shapes.length-1\)\].attr\(\{'opacity':0\}\);preso.animations.push\(\{'ids':\[\(preso.shapes.length-1\)\],'dur':" + anim.effect.Timing.Duration * 1000
                + @",'delay':" + anim.effect.Timing.TriggerDelayTime * 1000 + @",animate:\{'opacity':1\}\}\);";
            actual = target.toRaphJS();
            doesEqual = Regex.IsMatch(actual, expected);
            Assert.AreEqual(true, doesEqual);
        }
    }
}

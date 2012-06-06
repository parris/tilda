using Tilda.Models;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace TildaTests
{
    
    
    /// <summary>
    ///This is a test class for TildaAudioTest and is intended
    ///to contain all TildaAudioTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TildaAudioTest {


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
        ///A test for TildaAudio Constructor
        ///</summary>
        [TestMethod()]
        public void TildaAudioConstructorTest() {
            PowerPoint.Shape shape = new TildaTests.Mocks.MockShape(); 
            int id = 99; 
            TildaAudio target = new TildaAudio(shape, id);
            Assert.AreEqual(id, target.id);
            Assert.AreEqual(shape, target.shape);
        }

        /// <summary>
        /// Export audio file test
        /// </summary>
        [TestMethod()]
        public void exportAudioFile() {
            PowerPoint.Shape shape = new TildaTests.Mocks.MockShape(MsoShapeType.msoMedia);
            int id = 99;
            TildaAudio ta = new TildaAudio(shape, id);
            String actual = ta.toRaphJS();
            Assert.AreEqual("", actual);
        }
    }
}

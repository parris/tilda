using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {
    class MockParagraphFormat2 : ParagraphFormat2 {
        private MsoParagraphAlignment alignment = MsoParagraphAlignment.msoAlignLeft;

        public MsoParagraphAlignment Alignment {
            get {
                return this.alignment;
            }
            set {
                this.alignment = value;
            }
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public MsoBaselineAlignment BaselineAlignment {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public BulletFormat2 Bullet {
            get { throw new NotImplementedException(); }
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState FarEastLineBreakLevel {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float FirstLineIndent {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState HangingPunctuation {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int IndentLevel {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float LeftIndent {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState LineRuleAfter {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState LineRuleBefore {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState LineRuleWithin {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public float RightIndent {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float SpaceAfter {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float SpaceBefore {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float SpaceWithin {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public TabStops2 TabStops {
            get { throw new NotImplementedException(); }
        }

        public MsoTextDirection TextDirection {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState WordWrap {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
    }
}

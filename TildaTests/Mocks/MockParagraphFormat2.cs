using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {

    [Serializable]
    class MockParagraphFormat2 : ParagraphFormat2 {
        private MsoParagraphAlignment alignment = MsoParagraphAlignment.msoAlignLeft;
        private BulletFormat2 bullForm = new MockBulletFormat2();
        private float spaceWithin = .5f;
        private float spaceAfter = .5f;
        private float spaceBefore = .5f;
        private int indentLevel = 1;
        private float levelIndent = 8f;
        private float firstLineIndent = 15f;

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
            get { return this.bullForm; }
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
                return this.firstLineIndent;
            }
            set {
                this.firstLineIndent = value;
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
                return this.indentLevel;
            }
            set {
                this.indentLevel = value;
            }
        }

        public float LeftIndent {
            get {
                return this.levelIndent;
            }
            set {
                this.levelIndent = value;
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
                return this.spaceAfter;
            }
            set {
                this.spaceAfter = value;
            }
        }

        public float SpaceBefore {
            get {
                return this.spaceBefore;
            }
            set {
                this.spaceBefore = value;
            }
        }

        public float SpaceWithin {
            get {
                return this.spaceWithin;
            }
            set {
                this.spaceWithin = value;
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

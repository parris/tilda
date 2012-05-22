using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {

    [Serializable]
    class MockTextFrame2 : PowerPoint.TextFrame2 {
        private TextRange2 textRange;
        private float marginTop = 0f;
        private float marginBottom = 0f;
        private float marginLeft = 0f;
        private float marginRight = 0f;
        private MsoVerticalAnchor anchor = MsoVerticalAnchor.msoAnchorTop;

        public MockTextFrame2(TextRange2 tr = null) {
            if (tr == null)
                textRange = new MockTextRange2("",this);
            else
                textRange = tr;
        }

        public float MarginBottom {
            get {
                return this.marginBottom;
            }
            set {
                this.marginBottom = value;
            }
        }

        public float MarginLeft {
            get {
                return this.marginLeft;
            }
            set {
                this.marginLeft = value;
            }
        }

        public float MarginRight {
            get {
                return this.marginRight;
            }
            set {
                this.marginRight = value;
            }
        }

        public float MarginTop {
            get {
                return this.marginTop;
            }
            set {
                this.marginTop = value;
            }
        }

        public TextRange2 TextRange {
            get { return this.textRange; }
        }

        public MsoVerticalAnchor VerticalAnchor {
            get {
                return this.anchor;
            }
            set {
                this.anchor = value;
            }
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public MsoAutoSize AutoSize {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public TextColumn2 Column {
            get { throw new NotImplementedException(); }
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public void DeleteText() {
            throw new NotImplementedException();
        }

        public MsoTriState HasText {
            get { throw new NotImplementedException(); }
        }

        public MsoHorizontalAnchor HorizontalAnchor {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState NoTextRotation {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTextOrientation Orientation {
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

        public MsoPathFormat PathFormat {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Ruler2 Ruler {
            get { throw new NotImplementedException(); }
        }

        public PowerPoint.ThreeDFormat ThreeD {
            get { throw new NotImplementedException(); }
        }

        public MsoWarpFormat WarpFormat {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoPresetTextEffect WordArtFormat {
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

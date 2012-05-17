using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {
    class MockFont : Font2 {

        private String name = "Arial";
        private float size = 0f;
        private Office.FillFormat fill = new MockFillFormat();
        private MsoTriState bold = MsoTriState.msoFalse;
        private MsoTriState italic = MsoTriState.msoFalse;

        public MsoTriState Allcaps {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState AutorotateNumbers {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float BaselineOffset {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState Bold {
            get {
                return this.bold;
            }
            set {
                this.bold = value;
            }
        }

        public MsoTextCaps Caps {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState DoubleStrikeThrough {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState Embeddable {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Embedded {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Equalize {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public FillFormat Fill {
            get { return this.fill; }
        }

        public GlowFormat Glow {
            get { throw new NotImplementedException(); }
        }

        public ColorFormat Highlight {
            get { throw new NotImplementedException(); }
        }

        public MsoTriState Italic {
            get {
                return this.italic;
            }
            set {
                this.italic = value;
            }
        }

        public float Kerning {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public LineFormat Line {
            get { throw new NotImplementedException(); }
        }

        public string Name {
            get {
                return this.name;
            }
            set {
                this.name = value;
            }
        }

        public string NameAscii {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string NameComplexScript {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string NameFarEast {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string NameOther {
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

        public ReflectionFormat Reflection {
            get { throw new NotImplementedException(); }
        }

        public ShadowFormat Shadow {
            get { throw new NotImplementedException(); }
        }

        public float Size {
            get {
                return this.size;
            }
            set {
                this.size = value;
            }
        }

        public MsoTriState Smallcaps {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoSoftEdgeType SoftEdgeFormat {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float Spacing {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTextStrike Strike {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState StrikeThrough {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState Subscript {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState Superscript {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public ColorFormat UnderlineColor {
            get { throw new NotImplementedException(); }
        }

        public MsoTextUnderlineType UnderlineStyle {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoPresetTextEffect WordArtformat {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
    }
}

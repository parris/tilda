using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {

    [Serializable]
    class MockBulletFormat2 : BulletFormat2  {
        private MsoBulletType type = MsoBulletType.msoBulletNone;
        private Font2 font;
        private int character = 167;
        private float relativeSize = 2;

        public MockBulletFormat2(Font2 font = null) {
            if (font==null)
                this.font  = new MockFont();
            else
                this.font = font;
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public int Character {
            get {
                return this.character;
            }
            set {
                this.character = value;
            }
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public Font2 Font {
            get { return this.font; }
        }

        public int Number {
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public void Picture(string FileName) {
            throw new NotImplementedException();
        }

        public float RelativeSize {
            get {
                return this.relativeSize;
            }
            set {
                this.relativeSize = value;
            }
        }

        public int StartValue {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoNumberedBulletStyle Style {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoBulletType Type {
            get {
                return this.type;
            }
            set {
                this.type = value;
            }
        }

        public MsoTriState UseTextColor {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState UseTextFont {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoTriState Visible {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
    }
}

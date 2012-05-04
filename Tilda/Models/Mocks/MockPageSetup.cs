using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace Tilda.Models {
    class MockPageSetup : PageSetup {

        private float height = 1900;
        private float width = 1080;

        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public int FirstSlideNumber {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoOrientation NotesOrientation {
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

        public float SlideHeight {
            get {
                return height;
            }
            set {
                height = value;
            }
        }

        public Microsoft.Office.Core.MsoOrientation SlideOrientation {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public PpSlideSizeType SlideSize {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float SlideWidth {
            get {
                return width;
            }
            set {
                width = value;
            }
        }
    }
}

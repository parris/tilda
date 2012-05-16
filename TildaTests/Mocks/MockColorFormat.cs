using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockColorFormat : Microsoft.Office.Core.ColorFormat {

        private int rgb;
        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public float Brightness {
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

        public MsoThemeColorIndex ObjectThemeColor {
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

        public int RGB {
            get {
                return rgb;
            }
            set {
                this.rgb = value;
            }
        }

        public int SchemeColor {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float TintAndShade {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoColorType Type {
            get { throw new NotImplementedException(); }
        }
    }
}

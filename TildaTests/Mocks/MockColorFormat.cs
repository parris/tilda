using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockColorFormat : ColorFormat
    {
        private int rgb;

        public dynamic Application
        {
            get { throw new NotImplementedException(); }
        }

        public float Brightness
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public int Creator
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoThemeColorIndex ObjectThemeColor
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }

        public int RGB
        {
            get
            {
                return rgb;
            }
            set
            {
                this.rgb = value;
            }
        }

        public PpColorSchemeIndex SchemeColor
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public float TintAndShade
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoColorType Type
        {
            get { throw new NotImplementedException(); }
        }
    }
}

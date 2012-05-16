using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks
{
    public class MockTextEffectFormat : TextEffectFormat
    {
        private String fontName = "";
        private float fontSize = 16;
        private String text = "";

        public Microsoft.Office.Core.MsoTextEffectAlignment Alignment
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

        public dynamic Application
        {
            get { throw new NotImplementedException(); }
        }

        public int Creator
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState FontBold
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

        public Microsoft.Office.Core.MsoTriState FontItalic
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

        public string FontName
        {
            get { return this.fontName; }
            set { this.fontName = value; }
        }

        public float FontSize
        {
            get
            {
                return this.fontSize;
            }
            set
            {
                this.fontSize = value;
            }
        }

        public Microsoft.Office.Core.MsoTriState KernedPairs
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

        public Microsoft.Office.Core.MsoTriState NormalizedHeight
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

        public Microsoft.Office.Core.MsoPresetTextEffectShape PresetShape
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

        public Microsoft.Office.Core.MsoPresetTextEffect PresetTextEffect
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

        public Microsoft.Office.Core.MsoTriState RotatedChars
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

        public string Text
        {
            get
            {
                return text;
            }
            set
            {
                this.text = value;
            }
        }

        public void ToggleVerticalText()
        {
            throw new NotImplementedException();
        }

        public float Tracking
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
    }
}

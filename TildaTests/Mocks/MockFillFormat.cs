using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks
{
    class MockFillFormat : FillFormat
    {
        private ColorFormat backColor;
        private ColorFormat foreColor;

        public MockFillFormat()
        {
            this.backColor = new MockColorFormat();
            this.foreColor = new MockColorFormat();
        }

        public dynamic Application
        {
            get { throw new NotImplementedException(); }
        }

        public ColorFormat BackColor
        {
            get
            {
                return this.backColor;
            }
            set
            {
                this.backColor = value;
            }
        }

        public void Background()
        {
            throw new NotImplementedException();
        }

        public int Creator
        {
            get { throw new NotImplementedException(); }
        }

        public ColorFormat ForeColor
        {
            get
            {
                return this.foreColor;
            }
            set
            {
                this.foreColor = value;
            }
        }

        public float GradientAngle
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

        public Microsoft.Office.Core.MsoGradientColorType GradientColorType
        {
            get { throw new NotImplementedException(); }
        }

        public float GradientDegree
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.GradientStops GradientStops
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoGradientStyle GradientStyle
        {
            get { throw new NotImplementedException(); }
        }

        public int GradientVariant
        {
            get { throw new NotImplementedException(); }
        }

        public void OneColorGradient(Microsoft.Office.Core.MsoGradientStyle Style, int Variant, float Degree)
        {
            throw new NotImplementedException();
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoPatternType Pattern
        {
            get { throw new NotImplementedException(); }
        }

        public void Patterned(Microsoft.Office.Core.MsoPatternType Pattern)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.PictureEffects PictureEffects
        {
            get { throw new NotImplementedException(); }
        }

        public void PresetGradient(Microsoft.Office.Core.MsoGradientStyle Style, int Variant, Microsoft.Office.Core.MsoPresetGradientType PresetGradientType)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoPresetGradientType PresetGradientType
        {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoPresetTexture PresetTexture
        {
            get { throw new NotImplementedException(); }
        }

        public void PresetTextured(Microsoft.Office.Core.MsoPresetTexture PresetTexture)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState RotateWithObject
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

        public void Solid()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTextureAlignment TextureAlignment
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

        public float TextureHorizontalScale
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

        public string TextureName
        {
            get { throw new NotImplementedException(); }
        }

        public float TextureOffsetX
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

        public float TextureOffsetY
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

        public Microsoft.Office.Core.MsoTriState TextureTile
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

        public Microsoft.Office.Core.MsoTextureType TextureType
        {
            get { throw new NotImplementedException(); }
        }

        public float TextureVerticalScale
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

        public float Transparency
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

        public void TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle Style, int Variant)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoFillType Type
        {
            get { throw new NotImplementedException(); }
        }

        public void UserPicture(string PictureFile)
        {
            throw new NotImplementedException();
        }

        public void UserTextured(string TextureFile)
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState Visible
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

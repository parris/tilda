using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Drawing;

namespace Tilda.Models {

    class TildaPicture : TildaShape {


        /**
         * Creates a new TildaShape Object from a powerpoint shape
         * @param PowerPoint.Shape
         */
        public TildaPicture(PowerPoint.Shape shape, int id = 0)
            : base(shape, id) {
        }

        public double findX() {
            float value = 0f;
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                value = scaler * (shape.Width / 2 + shape.TextFrame.MarginLeft + shape.Left);
            else
                value = scaler * (shape.Left + shape.TextFrame.MarginLeft);
            return Math.Round(value);
        }

        /**
         * Find Vertical positioning
         */
        public double findY() {
            //vert positioning
            float value = 0f;
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                value = scaler * (shape.Height / 2 + shape.TextFrame.MarginTop + shape.Top);
            else if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignJustifyLow)
                value = scaler * (shape.Height - shape.TextFrame.MarginTop + shape.Top);
            else
                value = scaler * (shape.Top + shape.TextFrame.MarginTop);

            return Math.Round(value);
        }

        public override string toRaphJS(TildaAnimation[] animationMap, TildaSlide slide) {
            return "";
        }
    }
}

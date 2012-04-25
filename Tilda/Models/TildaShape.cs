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

namespace Tilda.Models
{
    class TildaShape
    {
        public PowerPoint.Shape shape;
        public float scaler = 1.4F;
        public float padding = 5F;
        public int id = 0;

        /**
         * Creates a new TildaShape Object from a powerpoint shape
         * @param PowerPoint.Shape
         */
        public TildaShape(PowerPoint.Shape shape, int id = 0)
        {
            this.shape = shape;
            this.id = id;
        }

        public String position(float xOffset = 0, float yOffset = 0)
        {
            return (this.findX()+xOffset) + "," + (this.findY()+yOffset);
        }

        /**
         * Find horizontal positioning
         */
        public double findX()
        {
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
        public double findY()
        {
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

        /**
         * Converts Office Color Integers to Hex Values
         * @param int rgb representing the color
         * @return string hex representing the color, starting with '#' character
         * @throws ArgumenetOutOfRangeException if red, green or blue is calculated to be greater than 255.
         */
        public String rgbToHex(int rgb)
        {
            int blue = (rgb/65536);
            int green = ((rgb - (65536 * blue)) / 256);
            int red = rgb - ((blue * 65536) + (green * 256));

            if (red > 255 || green > 255 || blue > 255)
                throw new ArgumentOutOfRangeException();

            String a = giveHex(Convert.ToInt32(Math.Floor((double)red / 16)));
            String b = giveHex(Convert.ToInt32(Math.Floor((double)red % 16)));
            String c = giveHex(Convert.ToInt32(Math.Floor((double)green / 16)));
            String d = giveHex(Convert.ToInt32(Math.Floor((double)green % 16)));
            String e = giveHex(Convert.ToInt32(Math.Floor((double)blue / 16)));
            String f = giveHex(Convert.ToInt32(Math.Floor((double)blue % 16)));

            String z = a + b + c + d + e + f;
            return "#"+z;
        }

        public String transformation()
        {
            float deg = shape.Rotation;
            return "'transformation':'r" + deg + "'";
        }

        private String giveHex(int dec)
        {
           if(dec == 10)
              return "a";
           else if(dec == 11)
              return "b";
           else if(dec == 12)
              return "c";
           else if(dec == 13)
              return "d";
           else if(dec == 14)
              return "e";
           else if(dec == 15)
              return "f";
           return ""+dec;
        }

        public String name()
        {
            if (shape.Type.Equals(Office.MsoShapeType.msoAutoShape))
                return "autoshape" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoCallout))
                return "callout" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoCanvas))
                return "canvas" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoChart))
                return "chart" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoComment))
                return "comment" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoDiagram))
                return "diagram" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoEmbeddedOLEObject))
                return "embeddedoleobj" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoFormControl))
                return "formcontrol" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoFreeform))
                return "freeform" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoGroup))
                return "group" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoInk))
                return "ink" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoInkComment))
                return "inkcomment" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoLine))
                return "line" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoLinkedOLEObject))
                return "linkedoleobj" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoLinkedPicture))
                return "linkedpicture" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoMedia))
                return "media" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoOLEControlObject))
                return "olecontrolobject" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoPicture))
                return "picture" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoPlaceholder))
                return "placeholder" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoScriptAnchor))
                return "scriptanchor" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoShapeTypeMixed))
                return "shapetypemixed" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoSlicer))
                return "slicer" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoSmartArt))
                return "smartart" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoTable))
                return "table" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoTextBox))
                return "textbox" + shape.Id;
            else if (shape.Type.Equals(Office.MsoShapeType.msoTextEffect))
                return "texteffect" + shape.Id;
            else
                return "unknownshape" + shape.Id;
        }

        public String toRaphJS() {
            return "";
        }
    }
}

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
        public List<TildaAnimation> animations;
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
            this.scaler = Settings.Scaler();
            this.id = id;
            animations = new List<TildaAnimation>();
        }

        public virtual String position(float xOffset = 0, float yOffset = 0)
        {
            return (this.findX()+xOffset) + "," + (this.findY()+yOffset);
        }

        /**
         * Find horizontal positioning
         */
        public virtual double findX()
        {
            return this.shape.Left * scaler;
        }

        /**
         * Find Vertical positioning
         */
        public virtual double findY()
        {
            return this.shape.Top * scaler;
        }

        /**
         * Converts Office Color Integers to Hex Values
         * @param int rgb representing the color
         * @return string hex representing the color, starting with '#' character
         * @throws ArgumenetOutOfRangeException if red, green or blue is calculated to be greater than 255.
         */
        public String rgbToHex(int rgb)
        {
            if(rgb < 0) //safeguard from weird stuff
                return "#000000";
            int blue = (rgb & 255);
            int green = (rgb >> 8) & 255;
            int red = (rgb >> 16) & 255;

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

        /**
         * Returns the rotation for this shape
         * @return String containing the amount to string is rotated in degrees
         */
        public String transformation()
        {
            float deg = shape.Rotation;
            return "'transformation':'r" + deg + "'";
        }

        /**
         * Gives a hex value for some value 0-15
         * @param int number to convert
         * @return String hex value
         */
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

        /**
         * In the current implementation this is not needed; however, if we switch to a model where
         * we have javascript selectors then this would be valuable.
         * @return String the name of this shape
         */
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

        /**
         * Somewhat of an abstract definition of how a shape should output to raphael js
         * @return String JS that would render this shape
         */
        public virtual String toRaphJS() {
            return "";
        }
    }
}

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

    class TildaTextbox : TildaShape {
        
        public String text = "";
        /**
         * Creates a new WBLShape Object from a powerpoint shape
         * @param PowerPoint.Shape
         */
        public TildaTextbox(PowerPoint.Shape shape, int id = 0): base(shape, id)
        {
            this.text = this.tildifyText();
        }

        public String fontStyle()
        {
            return "'font-style':'" + shape.TextEffect.FontName + "','font-size':'" + scaler * shape.TextEffect.FontSize + "','fill':'" + this.rgbToHex(shape.TextFrame.TextRange.Font.Color.RGB) + "'";
        }

        public String fontPosition(float addx = 0, float addy = 0) {
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                return "'cx':" + (this.findX() + addx) + ", 'cy':'" + (this.findY() + addy) + "', 'text-anchor': 'middle'";
            else
                return "'x':" + (this.findX() + addx) + ", 'y':'" + (this.findY() + addy) + "', 'text-anchor': 'start'";
        }

        public String position(float xOffset = 0, float yOffset = 0) {
            return (this.findX() + xOffset) + "," + (this.findY() + yOffset);
        }

        /**
         * Find horizontal positioning
         */
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

        public String tildifyText()
        {
            var font = this.fontStyle();
            var deg = this.transformation();
            String text = "";
            //we choose to represent line breaks as "~|" to keep the same length and not interfere with anything
            foreach (TextRange paragraph in shape.TextFrame.TextRange.Paragraphs()) {
                var pgText = paragraph.Text.Replace("\r", "~|");
                var lines = paragraph.Lines(0,400);
                var pos = 0;
                var count = 0;
                if (lines.Count > 1) 
                    foreach (TextRange line in lines) {
                        pos += line.Length;
                        if (count < lines.Count-1)
                            pgText = pgText.Insert(pos, "~");
                        count++;
                    }
                if (paragraph.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone)
                    pgText = "-" + paragraph.IndentLevel + " " + pgText;
                text += pgText;
            }
            
            return text;
        }

        public static string toRaphJS(TildaShape[] shapeMap, TildaAnimation[] animationMap) {
            String js = "";
            int shapeCount = 0;
            foreach (TildaTextbox shape in shapeMap) {
                if (shape == null)
                    continue;
                double lineHeight = (shape.shape.TextFrame.TextRange.Font.Size + shape.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin) * shape.scaler;
                double currentHeight;
                if (shape.shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignLeft)
                    currentHeight = shape.findY() + (float)(lineHeight / (1.5));
                else
                    currentHeight = shape.findY();
                string font = shape.fontStyle();
                string transform = shape.transformation();
                double shapeX = shape.findX();
                String[] parts = shape.text.Split(new string[] { "~|" }, StringSplitOptions.None);
                for (int i = 0; i < parts.Length; i++) {
                    String part = parts[i];
                    TildaAnimation found = null;
                    int shapeAnim = -1;
                    string textboxAnims = "";
                    //find animation
                    foreach (TildaAnimation animation in animationMap)
                        if (found == null && shape.shape.Id.Equals(animation.shape.shape.Id) && i == animation.effect.Paragraph - 1)
                            found = animation;

                    //is bullet? add some spacing...
                    double xAdd = 0;
                    bool hasBullet = false;

                    if (part.Length > 0 && part[0] == '-') {
                        hasBullet = true;
                        float bulletSize = shape.shape.TextFrame.TextRange.Font.Size / 4 * shape.scaler;
                        js += "shapes.push(paper.rect(" + (shapeX + 5) + "," + (currentHeight - bulletSize / 2) + "," + bulletSize + "," + bulletSize + ").attr({'stroke':'#84BD00','fill':'#84BD00'}));";
                        if (found != null) {
                            js += "shapes[" + shapeCount + "].attr({'fill-opacity':0,'stroke-opacity':0});";
                            shapeAnim = shapeCount;
                        }
                        xAdd += 30 * shape.scaler;
                        part = part.Substring(3);
                        shapeCount++;
                    }

                    //split even more
                    String[] miniparts = part.Split('~');
                    foreach (String minipart in miniparts) {
                        var fontpos = shape.fontPosition((float)xAdd, (float)(currentHeight - shape.findY()));
                        String textbox = "shape.push(paper.text(" + (shapeX + xAdd) + "," + currentHeight + ",'" + minipart + "').attr({" + font + "," + transform + "," + fontpos + "}));";

                        if (found != null) {
                            textbox += "shape[" + shapeCount + "].attr({'fill-opacity':0,'stroke-opacity':0});";
                            textboxAnims += shapeCount + ",";
                        }

                        js += textbox;
                        currentHeight += lineHeight;
                        shapeCount++;
                    }

                    //more bullet stuff
                    if (hasBullet)
                        currentHeight += lineHeight / 3; // some extra amount

                    if (textboxAnims.Length > 0)
                        js += "animations.push({'ids':[" + textboxAnims.Substring(1) + shapeAnim +"],'dur':" + found.effect.Timing.Duration * 1000 + ",'delay':" + found.effect.Timing.TriggerDelayTime * 1000 + "});";
                }
            }
            return js;
        }
    }
}

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
         * Creates a new TildaShape Object from a powerpoint shape
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
            else if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignRight)
                return "'x':" + (this.findX() + addx) + ", 'y':'" + (this.findY() + addy) + "', 'text-anchor': 'end'";
            else
                return "'x':" + (this.findX() + addx) + ", 'y':'" + (this.findY() + addy) + "', 'text-anchor': 'start'";
        }

        public override String position(float xOffset = 0, float yOffset = 0) {
            return (this.findX() + xOffset) + "," + (this.findY() + yOffset);
        }

        /**
         * Find horizontal positioning
         */
        public override double findX() {
            float value = 0f;
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                value = scaler * (shape.Width / 2 + shape.TextFrame.MarginLeft + shape.Left);
            else if(shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignRight)
                value = scaler * (shape.Left + shape.Width - shape.TextFrame.MarginRight);
            else
                value = scaler * (shape.Left + shape.TextFrame.MarginLeft);
            return Math.Round(value);
        }

        /**
         * Find Vertical positioning
         */
        public override double findY() {
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

        public override string toRaphJS(TildaAnimation[] animationMap,TildaSlide slide) {
            String js = "";
            double lineHeight = (shape.TextFrame.TextRange.Font.Size + this.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin) * this.scaler;
            double currentHeight;
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignLeft)
                currentHeight = this.findY() + (float)(lineHeight / (1.5));
            else
                currentHeight = this.findY();
            string font = this.fontStyle();
            string transform = this.transformation();
            double shapeX = this.findX();
            String[] parts = this.text.Split(new string[] { "~|" }, StringSplitOptions.None);
            for (int i = 0; i < parts.Length; i++) {
                String part = parts[i];
                TildaAnimation found = null;
                int shapeAnim = -1;
                string textboxAnims = "";
                //find animation
                foreach(TildaAnimation animation in animationMap) {
                    try {
                        if(found == null && this.shape.Id.Equals(animation.shape.shape.Id) && i == animation.effect.Paragraph - 1)
                            found = animation;
                    } catch(Exception e) { } // this is obviously not the animation we are looking for; however, just throw it away rather than complaining
                }

                //is bullet? add some spacing...
                double xAdd = 0;
                bool hasBullet = false;

                if (part.Length > 0 && part[0] == '-') {
                    currentHeight += (lineHeight / 3)/2; // some extra amount,before
                    hasBullet = true;
                    float bulletSize = this.shape.TextFrame.TextRange.Font.Size / 4 * this.scaler;
                    //if bullet is not at level 1 figure out how much space to add
                    int bulletXSpace = 0;
                    if(Int32.Parse(part[1] + "") > 1)
                        bulletXSpace += 30 * Int32.Parse(part[1] + "");
                    js += "preso.shapes.push(preso.paper.rect(" + (shapeX + 5 + bulletXSpace) + "," + (currentHeight - bulletSize / 2) + "," + bulletSize + "," + bulletSize + ").attr({'stroke':'#84BD00','fill':'#84BD00'}));";
                    if (found != null) {
                        js += "preso.shapes[" + slide.shapeCount + "].attr({'fill-opacity':0,'stroke-opacity':0});";
                        shapeAnim = slide.shapeCount;
                    }
                    xAdd += (30 * this.scaler) * Int32.Parse(""+part[1]);
                    part = part.Substring(3);
                    slide.shapeCount++;
                }

                //split even more
                String[] miniparts = part.Split('~');
                foreach (String minipart in miniparts) {
                    var fontpos = this.fontPosition((float)xAdd, (float)(currentHeight - this.findY()));
                    String textbox = "preso.shapes.push(preso.paper.text(" + (shapeX + xAdd) + "," + currentHeight + ",'" + minipart + "').attr({" + font + "," + transform + "," + fontpos + "}));";

                    if (found != null) {
                        textbox += "preso.shapes[" + slide.shapeCount + "].attr({'fill-opacity':0,'stroke-opacity':0,'opacity':0});";
                        textboxAnims += slide.shapeCount + ",";
                    }

                    js += textbox;
                    currentHeight += lineHeight;
                    slide.shapeCount++;
                }

                //more bullet stuff
                if (hasBullet)
                    currentHeight += (lineHeight / 3) / 2; // some extra amount,after

                if (textboxAnims.Length > 0) {
                    string ids = textboxAnims;
                    if (shapeAnim != -1)
                        ids += shapeAnim;
                    else
                        ids = ids.Substring(0, ids.Length - 1);
                    js += "preso.animations.push({'ids':[" + ids + "],'dur':" + found.effect.Timing.Duration * 1000 + ",'delay':" + found.effect.Timing.TriggerDelayTime * 1000 + ",animate:{'fill-opacity':1,'stroke-opacity':1,'opacity':1}});";
                }
            }
            return js;
        }
    }
}

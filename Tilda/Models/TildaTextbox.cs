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
        private List<TextRange> paragraphs = new List<TextRange>();

        /**
         * Creates a new TildaShape Object from a powerpoint shape
         * @param PowerPoint.Shape
         */
        public TildaTextbox(PowerPoint.Shape shape, int id = 0): base(shape, id)
        {
            //this.text = this.getParagraphs();
            this.getParagraphs();
        }

        public String fontStyle(TextRange range = null)
        {
            if(range == null)
                range = shape.TextFrame.TextRange;
            return "'font-style':'" + range.Font.Name + "','font-size':'" + scaler * range.Font.Size + "','fill':'" + this.rgbToHex(range.Font.Color.RGB) + "'";
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

        public override String toRaphJS() {
            String js = "";
            double lineHeight = (shape.TextFrame.TextRange.Font.Size + this.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin) * this.scaler;
            double currentHeight;
            if(shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignLeft)
                currentHeight = this.findY() + (float)(lineHeight / (1.5)); //this feels odd, but it looks right
            else
                currentHeight = this.findY();
            double shapeX = this.findX();

            string font = this.fontStyle();
            string transform = this.transformation();

            for (int i = 0; i < paragraphs.Count; i++) {
                js += "idsToAnimate = new Array();";
                TextRange paragraph = paragraphs.ElementAt(i);
                TildaAnimation found = null;
                //find animation
                foreach(TildaAnimation animation in animations) {
                    try {
                        if(found == null && this.shape.Id.Equals(animation.shape.shape.Id) && i == animation.effect.Paragraph - 1)
                            found = animation;
                    } catch(Exception e) { } // this is obviously not the animation we are looking for; however, just throw it away rather than complaining
                }

                //is bullet? add some spacing...
                double xAdd = 0;
                bool hasBullet = false;
                if(paragraph.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone) {
                    currentHeight += (lineHeight / 3) / 2; // some extra amount,before
                    hasBullet = true;
                    float bulletSize = paragraph.Font.Size / 4 * this.scaler;
                    int bulletXSpace = 30 * paragraph.IndentLevel;
                    js += "preso.shapes.push(preso.paper.rect(" + (shapeX + 5 + bulletXSpace) + "," + (currentHeight - bulletSize / 2) + "," + bulletSize + "," + bulletSize + ").attr({'stroke':'#84BD00','fill':'#84BD00'}));";
                    if(found != null) {
                        js += "idsToAnimate.push(preso.shapes.length-1);";
                        js += "preso.shapes[(preso.shapes.length-1)].attr({'fill-opacity':0,'stroke-opacity':0});";
                    }
                    xAdd += bulletXSpace*this.scaler;
                }

                var lines = paragraph.Lines(0, 400);
                foreach(TextRange minipart in lines) {
                    var fontpos = this.fontPosition((float)xAdd, (float)(currentHeight - this.findY()));
                    String textbox = "preso.shapes.push(preso.paper.text(" + (shapeX + xAdd) + "," + currentHeight + ",'" + minipart.Text.Replace("\r", "") +"').attr({" + font + "," + transform + "," + fontpos + "}));";

                    if(found != null) {
                        textbox += "idsToAnimate.push(preso.shapes.length-1);";
                        textbox += "preso.shapes[(preso.shapes.length-1)].attr({'fill-opacity':0,'stroke-opacity':0,'opacity':0});";
                        //textboxAnims += slide.shapeCount + ",";
                    }

                    js += textbox;
                    currentHeight += lineHeight;
                }

                //extra bullet spacing
                if(hasBullet)
                    currentHeight += (lineHeight / 3) / 2; // some extra amount,after

                if(found != null)
                    js += "preso.animations.push({'ids':idsToAnimate,'dur':" + found.effect.Timing.Duration * 1000 + ",'delay':" + found.effect.Timing.TriggerDelayTime * 1000 + ",animate:{'fill-opacity':1,'stroke-opacity':1,'opacity':1}});";
            }

            return js;
        }

        private void getParagraphs() {
            foreach(TextRange paragraph in shape.TextFrame.TextRange.Paragraphs()) {
                this.paragraphs.Add(paragraph);
            }
        }
    }
}

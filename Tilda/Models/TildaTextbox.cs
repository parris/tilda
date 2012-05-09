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
        //used when calculating the height
        //reduces the amount of code, although is modified and read in a few functions
        private double currentHeight = 0; 

        /**
         * Creates a new TildaShape Object from a powerpoint shape
         * @param PowerPoint.Shape
         */
        public TildaTextbox(PowerPoint.Shape shape, int id = 0): base(shape, id)
        {
            //this.text = this.getParagraphs();
            this.getParagraphs();
        }

        /**
         * JSON Object attributes for font style of a line of text
         * @param TextRange to look up information on, if none specifed then this this.shape's TextRange is found and used
         * @return String representing JSON attributes of font style as per RaphaelJS specifications
         */
        public String fontStyle(TextRange range = null)
        {
            if(range == null)
                range = shape.TextFrame.TextRange;
            return "'font-style':'" + range.Font.Name + "','font-size':'" + scaler * range.Font.Size + "','fill':'" + this.rgbToHex(range.Font.Color.RGB) + "'";
        }

        /**
         * Part of a JSON object representing the font position for RaphJS. Based on SVG and VML requirements
         * @param float amount to offset x position by.
         * @param float amount to offset y position by.
         * @return String representing part of JSON object for text attributes in RaphJS
         */
        public String fontPosition(float addx = 0, float addy = 0) {
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                return "'cx':" + (this.findX() + addx) + ", 'cy':'" + (this.findY() + addy) + "', 'text-anchor': 'middle'";
            else if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignRight)
                return "'x':" + (this.findX() + addx) + ", 'y':'" + (this.findY() + addy) + "', 'text-anchor': 'end'";
            else
                return "'x':" + (this.findX() + addx) + ", 'y':'" + (this.findY() + addy) + "', 'text-anchor': 'start'";
        }

        /**
         * X,Y coordinates, comma seperated for RaphJS functions
         * @param float amount to offset x position by.
         * @param float amount to offset y position by.
         * @return String representing coordinates, comma seperated.
         */
        public override String position(float xOffset = 0, float yOffset = 0) {
            return (this.findX() + xOffset) + "," + (this.findY() + yOffset);
        }

        /**
         * Find horizontal positioning
         * @return doulbe x position of this Shape for RaphJS purposes
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
         * @return double y position of this Shape for RaphJS purposes
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

        /**
         * Will render this Text shape to RaphJS code
         * @see currentHeight (private member) will be modified and read by this method.
         * @see PowerPoint.Shape, PowerPoint.TextFrame, PowerPoint.TextRange for more information about 
         * what this method expects. 
         * @see Settings.Scaler() 
         * @return String representing RaphJS code
         */
        public override String toRaphJS() {
            this.currentHeight = 0;
            String js = "";
            double lineHeight = (shape.TextFrame.TextRange.Font.Size + this.shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin) * this.scaler;
            if(shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignLeft)
                this.currentHeight = this.findY() + (float)(lineHeight / 4); //this feels odd, but it "looks" right
            else
                this.currentHeight = this.findY();
            double shapeX = this.findX();

            for (int i = 0; i < paragraphs.Count; i++) {
                js += "idsToAnimate = new Array();";
                TextRange paragraph = paragraphs.ElementAt(i);
                TildaAnimation found = null;
                //find animation
                foreach(TildaAnimation animation in animations) {
                    try {
                        if(found == null && this.shape.Id.Equals(animation.shape.shape.Id) && i == animation.effect.Paragraph - 1)
                            found = animation;
                    } catch{ } // this is obviously not the animation we are looking for; however, just throw it away rather than complaining
                }

                //add spacing above this line
                this.currentHeight += (paragraph.ParagraphFormat.SpaceBefore) * this.scaler;

                double xAdd = 0;
                if(paragraph.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone) {
                    float[] offsets = this.findIndentSpacing(paragraph,paragraph.IndentLevel);
                    float bulletXSpace = offsets[0];
                    float bulletSize = paragraph.Font.Size / 4 * this.scaler;

                    js += this.renderBullet(paragraph, (shapeX + 5 + bulletXSpace), (this.currentHeight - bulletSize / 2), found);
                    xAdd += offsets[1]+bulletSize;
                }

                foreach(TextRange line in paragraph.Lines(0, 400)) 
                    js += this.renderLine(line,xAdd,found);

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

        /**
         * Renders a Line of Text to RaphJS code
         * @see currentHeight (private member) will be modified and read by this method.
         * @param TextRange that represents the line of text
         * @param double x position offset if needed
         * @param TildaAnimation object containing a powerpoint animation, or just null if no animation
         * @return String containing the RaphJS code representing this bullet
         */
        private String renderLine(TextRange line, double xOffset, TildaAnimation anim = null) {
            //All text seems 1 letter off approximately in the x axis, so let's just shift about 1 letter over
            //I will estimate this at 1/8 the size of the lineheight
            xOffset -= (line.Font.Size*this.scaler) / 8;

            string font = this.fontStyle(line);
            string transform = this.transformation();
            var fontpos = this.fontPosition((float)(xOffset), (float)(this.currentHeight - this.findY()));

            String textbox = "preso.shapes.push(preso.paper.text(" + (this.findX() + xOffset) + "," + this.currentHeight + ",'" + line.Text.Replace("\r", "") + "').attr({" + font + "," + transform + "," + fontpos + "}));";

            if(anim != null) {
                textbox += "idsToAnimate.push(preso.shapes.length-1);";
                textbox += "preso.shapes[(preso.shapes.length-1)].attr({'fill-opacity':0,'stroke-opacity':0,'opacity':0});";
            }

            //push the current height position down 
            this.currentHeight += (line.Font.Size + line.ParagraphFormat.SpaceAfter) * this.scaler;

            //add some extra spacing, I tried to be mathematical about it, but I don't know where it is coming from
            //it looks like 1/4 the line height, but spread above and below it, the other half of this is before the bullet
            if(line.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone)
                this.currentHeight += ((line.Font.Size*this.scaler) / 8);
            return textbox;
        }

        /**
         * Tries to find the actual level of indentation between bullets and text at each level.
         * Currently a function is defined to calculate approximately how this should look. 
         * @param TextRange to look up indentation levels on.
         * @param int the level of the bullet for which information should be return
         * @return float[] contains 2 values FirstMargin (where the bullet is) and LeftMargin (where the text starts)
         */
        private float[] findIndentSpacing(TextRange t, int level) {
            float tabValue = 26f;
            /*if(level == 1) {
                RulerLevel rl = t.Parent.Ruler.Levels(2);
                //bullet must start at 0 on the first level for now
                return new float[2] { 0, rl.LeftMargin * this.scaler };
            } else {
                RulerLevel rl = t.Parent.Ruler.Levels[level];
                return new float[2] { rl.FirstMargin * this.scaler, rl.LeftMargin * this.scaler };
            }*/
            //it should look something like this:
            return new float[2] { (tabValue * (level - 1)) * this.scaler, (tabValue * (level)-level*2) * this.scaler };
        }

        /**
         * Looks up information about a bullet and creates RaphJS code to render it
         * @see currentHeight (private member) will be modified and read by this method.
         * @param TextRange the paragraph that has a bullet
         * @param double x position of the bullet
         * @param double y position of the bullet
         * @param TildaAnimation object containing a powerpoint animation, or just null if no animation
         * @return String containing the RaphJS code representing this bullet
         */
        private String renderBullet(TextRange t, double x, double y, TildaAnimation anim = null) {
            String js = "";

            //no bullet if no text
            if(t.Text == "" || t.Text == "\r")
                return js;

            //add some extra spacing, I tried to be mathematical about it, but I don't know where it is coming from
            //it looks like 1/3 the line height, but spread above and below it, the other half of this is after render line
            float extraSpacing = ((t.Font.Size*this.scaler) / 8);
            y += extraSpacing;
            this.currentHeight += extraSpacing;

            //relative size is set by user, this look approximately correct
            float bulletSize = (t.ParagraphFormat.Bullet.RelativeSize * (t.Font.Size / 4)) * this.scaler;
            int bullet = t.ParagraphFormat.Bullet.Character;

            // find the right color of the bullet, first try the color of the bullet itself
            // fall back to line color otherwise
            int rgb = t.ParagraphFormat.Bullet.Font.Color.RGB;
            if(rgb == 0)
                rgb = t.Font.Color.RGB;
            String color = this.rgbToHex(rgb);

            if(t.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletNumbered) {
                int bulletNumber = (t.ParagraphFormat.Bullet.StartValue - 1 + t.ParagraphFormat.Bullet.Number);
                String bulletText = this.numberedBullet(bulletNumber, t.ParagraphFormat.Bullet.Style);
                js += "preso.shapes.push(preso.paper.text(" + (x+bulletSize) + "," + this.currentHeight + ",'" + bulletText + "'" + ").attr({" + this.fontStyle(t) + "})";
            } else if(bullet == 8226) {
                double radius = bulletSize / 2;
                x += radius;
                y += radius;
                js += "preso.shapes.push(preso.paper.circle(" + x + "," + y + "," + radius + ")";
            } else // if(bullet == 167), the square
                js += "preso.shapes.push(preso.paper.rect(" + x + "," + y + "," + bulletSize + "," + bulletSize + ")";
            js += ".attr({'stroke':'" + color + "','fill':'" + color + "'}));";
            if(anim != null) {
                js += "idsToAnimate.push(preso.shapes.length-1);";
                js += "preso.shapes[(preso.shapes.length-1)].attr({'fill-opacity':0,'stroke-opacity':0});";
            }
            return js;
        }

        /**
         * Supports most common bullet styles. Should be extend to support anything potentially
         * @see settings for lists of values
         * @param Integer number of the bullet
         * @param PpNumberedBulletStyle from a numbered bullet type, from a textrange
         * @return String representing how the bullet should look
         */
        private String numberedBullet(int number,PpNumberedBulletStyle style) {
            String text = "";
            if(style == PpNumberedBulletStyle.ppBulletAlphaLCParenBoth)
                text = "(" + Settings.aToz[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletAlphaLCParenRight)
                text = Settings.aToz[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletAlphaLCPeriod)
                text = Settings.aToz[number] + ".";
            else if(style == PpNumberedBulletStyle.ppBulletAlphaUCParenBoth)
                text = "(" + Settings.AToZ[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletAlphaUCParenRight)
                text = Settings.AToZ[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletAlphaUCPeriod)
                text = Settings.AToZ[number] + ".";
            else if(style == PpNumberedBulletStyle.ppBulletArabicParenBoth)
                text = "("+ number + ")";
            else if(style == PpNumberedBulletStyle.ppBulletArabicParenRight)
                text = number + ")";
            else if(style == PpNumberedBulletStyle.ppBulletArabicPeriod || style == PpNumberedBulletStyle.ppBulletArabicDBPeriod)
                text = number + ".";
            else if(style == PpNumberedBulletStyle.ppBulletArabicPlain || style == PpNumberedBulletStyle.ppBulletArabicDBPlain)
                text = number + ".";
            else if(style == PpNumberedBulletStyle.ppBulletRomanLCParenBoth)
                text = "(" + Settings.romanNumeralsLC[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletRomanLCParenRight)
                text = Settings.romanNumeralsLC[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletRomanLCPeriod)
                text = Settings.romanNumeralsLC[number] + ".";
            else if(style == PpNumberedBulletStyle.ppBulletRomanUCParenBoth)
                text = "(" + Settings.romanNumeralsUC[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletRomanUCParenRight)
                text = Settings.romanNumeralsUC[number] + ")";
            else if(style == PpNumberedBulletStyle.ppBulletRomanUCPeriod)
                text = Settings.romanNumeralsUC[number] + ".";
            else
                text = number + ".";
            return text;
        }
    }
}

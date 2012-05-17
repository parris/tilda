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
using Microsoft.Office.Core;

namespace Tilda.Models {

    class TildaTextbox : TildaShape {
        
        public String text = "";
        private List<TextRange2> paragraphs = new List<TextRange2>();
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
        public String fontStyle(TextRange2 range = null)
        {
            if(range == null)
                range = shape.TextFrame2.TextRange;
            String js = "'font-size':'" + scaler * range.Font.Size + "','fill':'" + this.rgbToHex(range.Font.Fill.ForeColor.RGB) + "'";
            if(range.Font.Bold == MsoTriState.msoCTrue || range.Font.Bold == MsoTriState.msoTrue)
                js += ",'font-weight':'bold'";
            if(range.Font.Italic == MsoTriState.msoCTrue || range.Font.Italic == MsoTriState.msoTrue)
                js += ",'font-family':'" + range.Font.Name + " italic'";
            else
                js += ",'font-family':'" + range.Font.Name + "'";
            return js;
        }

        /**
         * Part of a JSON object representing the font position for RaphJS. Based on SVG and VML requirements
         * @param float amount to offset x position by.
         * @param float amount to offset y position by.
         * @return String representing part of JSON object for text attributes in RaphJS
         */
        public String fontPosition(float addx = 0, float addy = 0) {
            if (shape.TextFrame2.TextRange.ParagraphFormat.Alignment == MsoParagraphAlignment.msoAlignCenter)
                return "'text-anchor': 'middle'";
            else if(shape.TextFrame2.TextRange.ParagraphFormat.Alignment == MsoParagraphAlignment.msoAlignRight)
                return "'text-anchor': 'end'";
            else
                return "'text-anchor': 'start'";
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
         * Find horizontal positioning of this shape
         * @return doulbe x position of this Shape for RaphJS purposes
         */
        public override double findX() {
            float value = 0f;
            MsoParagraphAlignment alignment = shape.TextFrame2.TextRange.ParagraphFormat.Alignment;
            if(alignment == MsoParagraphAlignment.msoAlignCenter)
                value = scaler * (shape.Width / 2 + shape.TextFrame2.MarginLeft + shape.Left);
            else if(alignment == MsoParagraphAlignment.msoAlignRight)
                value = scaler * (shape.Left + shape.Width - shape.TextFrame2.MarginRight);
            else
                value = scaler * (shape.Left + shape.TextFrame2.MarginLeft);
            return value;
        }

        /**
         * Find vertical positioning of the starting point of the text of this shape.
         * If anchor top or middle the y position will be the top of line of the first line of text
         * If anchor bottom or baseline then the y position will be the bottom most portion of the text.
         * @return double y position of this Shape for RaphJS purposes
         */
        public override double findY() {
            //vert positioning
            float value = 0f;
            if (shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorMiddle)
                value = scaler * (shape.Height / 2 + shape.TextFrame2.MarginTop + shape.Top);
            else if(shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorBottom)
                value = scaler * (shape.Top + shape.Height - shape.TextFrame2.MarginBottom);
            else if(shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorBottomBaseLine)
                value = scaler * (shape.Top + shape.Height);
            else if(shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorTopBaseline)
                value = scaler * (shape.Top);
            else
                value = scaler * (shape.Top + shape.TextFrame2.MarginTop);

            return value;
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
            this.currentHeight = this.findY();
            double shapeX = this.findX();

            for (int i = 0; i < paragraphs.Count; i++) {
                js += "idsToAnimate = new Array();";
                TextRange2 paragraph = paragraphs.ElementAt(i);
                TildaAnimation found = null;
                //find animation
                foreach(TildaAnimation animation in animations) {
                    try {
                        if(found == null && this.shape.Id.Equals(animation.shape.shape.Id) && i == animation.effect.Paragraph - 1)
                            found = animation;
                    } catch{ } // this is obviously not the animation we are looking for; however, just throw it away rather than complaining
                }

                //add spacing above this line
                if (this.isBottomOrBaseLine())
                    this.currentHeight -= (paragraph.ParagraphFormat.SpaceAfter) * this.scaler;
                else
                    this.currentHeight += (paragraph.ParagraphFormat.SpaceBefore) * this.scaler;

                double xAdd = 0;
                if((PpBulletType)paragraph.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone) {
                    float[] offsets = this.findIndentSpacing(paragraph);
                    float bulletXSpace = offsets[0];
                    float bulletSize = paragraph.Font.Size / 4 * this.scaler; // seems correct

                    js += this.renderBullet(paragraph, (shapeX + bulletXSpace), (this.currentHeight - bulletSize / 2), found);
                    xAdd += offsets[1]+bulletSize;
                }

                foreach(TextRange2 line in this.getLines(paragraph.Lines)) 
                    js += this.renderLine(line,xAdd,found);

                if(found != null)
                    js += "preso.animations.push({'ids':idsToAnimate,'dur':" + found.effect.Timing.Duration * 1000 + ",'delay':" + found.effect.Timing.TriggerDelayTime * 1000 + ",animate:{'fill-opacity':1,'stroke-opacity':1,'opacity':1}});";
            }
            return js;
        }

        private List<TextRange2> getLines(TextRange2 lines) {
            List<TextRange2> trlines = new List<TextRange2>();
            foreach(TextRange2 line in lines)
                trlines.Add(line);

            if(this.isBottomOrBaseLine())
                trlines.Reverse();

            return trlines;
        }

        private void getParagraphs() {
            foreach(TextRange2 paragraph in shape.TextFrame2.TextRange.Paragraphs) {
                this.paragraphs.Add(paragraph);
            }

            if(this.isBottomOrBaseLine())
                this.paragraphs.Reverse();
        }

        private bool isBottomOrBaseLine() {
            if(shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorBottom
                || shape.TextFrame2.VerticalAnchor == MsoVerticalAnchor.msoAnchorBottomBaseLine)
                return true;
            
            return false;
        }

        /**
         * Renders a Line of Text to RaphJS code
         * @see currentHeight (private member) will be modified and read by this method.
         * @param TextRange that represents the line of text
         * @param double x position offset if needed
         * @param TildaAnimation object containing a powerpoint animation, or just null if no animation
         * @return String containing the RaphJS code representing this bullet
         */
        private String renderLine(TextRange2 line, double xOffset, TildaAnimation anim = null) {
            //push current height position up to write the text in the right place if going bottom up
            if(this.isBottomOrBaseLine())
                this.currentHeight -= (line.Font.Size * this.scaler);

            string font = this.fontStyle(line);
            string transform = this.transformation();
            var fontpos = this.fontPosition((float)(xOffset), (float)(this.currentHeight - this.findY()));

            String textbox = "preso.shapes.push(preso.paper.text(" + (this.findX() + xOffset) + "," + this.currentHeight + ",'" + line.Text.Replace("\r", "").Replace("\v", "") + "').attr({" + font + "," + transform + "," + fontpos + "}));";

            if(anim != null) {
                textbox += "idsToAnimate.push(preso.shapes.length-1);";
                textbox += "preso.shapes[(preso.shapes.length-1)].attr({'fill-opacity':0,'stroke-opacity':0,'opacity':0});";
            }

            //push the current height position
            if(this.isBottomOrBaseLine())
                this.currentHeight -= (line.ParagraphFormat.SpaceBefore) * this.scaler;
            else
                this.currentHeight += (line.Font.Size + line.ParagraphFormat.SpaceAfter) * this.scaler;


            //add some extra spacing, I tried to be mathematical about it, but I don't know where it is coming from
            //it looks like 1/4 the line height, but spread above and below it, the other half of this is before the bullet
            if((PpBulletType)line.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone)
                if (this.isBottomOrBaseLine())
                    this.currentHeight -= ((line.Font.Size * this.scaler) / 8);
                else
                    this.currentHeight += ((line.Font.Size*this.scaler) / 8);
            return textbox;
        }

        /**
         * Finds the level of indentation between bullets and text at each paragraph.
         * @param TextRange to look up indentation levels on.
         * @return float[] contains 2 values FirstMargin (where the bullet is) and LeftMargin (where the text starts)
         */
        private float[] findIndentSpacing(TextRange2 t) {
            ParagraphFormat2 pg = t.ParagraphFormat;
            return new float[2] { (pg.LeftIndent + pg.FirstLineIndent) * this.scaler, pg.LeftIndent * this.scaler };
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
        private String renderBullet(TextRange2 t, double x, double y, TildaAnimation anim = null) {
            String js = "";

            //no bullet if no text
            if(t.Text == "" || t.Text == "\r")
                return js;

            //add some extra spacing, I tried to be mathematical about it, but I don't know where it is coming from
            float extraSpacing = ((t.Font.Size*this.scaler) / 8);
            y += extraSpacing;
            if(this.isBottomOrBaseLine())
                this.currentHeight -= extraSpacing;
            else
                this.currentHeight += extraSpacing;

            //relative size is set by user, this look approximately correct
            float bulletSize = (t.ParagraphFormat.Bullet.RelativeSize * (t.Font.Size / 4)) * this.scaler;
            int bullet = t.ParagraphFormat.Bullet.Character;

            // find the right color of the bullet, first try the color of the bullet itself
            // fall back to line color otherwise
            int rgb = t.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB;
            if(rgb == 0)
                rgb = t.Font.Fill.ForeColor.RGB;
            String color = this.rgbToHex(rgb);

            if(t.ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletNumbered) {
                int bulletNumber = (t.ParagraphFormat.Bullet.StartValue - 1 + t.ParagraphFormat.Bullet.Number);
                String bulletText = this.numberedBullet(bulletNumber, t.ParagraphFormat.Bullet.Style);
                js += "preso.shapes.push(preso.paper.text(" + (x+bulletSize*2) + "," + this.currentHeight + ",'" + bulletText + "'" + ").attr({" + this.fontStyle(t) + "})";
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
        private String numberedBullet(int number,MsoNumberedBulletStyle style) {
            String text = "";
            if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaLCParenBoth)
                text = "(" + Settings.aToz[number-1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaLCParenRight)
                text = Settings.aToz[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaLCPeriod)
                text = Settings.aToz[number - 1] + ".";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaUCParenBoth)
                text = "(" + Settings.AToZ[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaUCParenRight)
                text = Settings.AToZ[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletAlphaUCPeriod)
                text = Settings.AToZ[number - 1] + ".";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicParenBoth)
                text = "("+ number + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicParenRight)
                text = number + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicPeriod || (PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicDBPeriod)
                text = number + ".";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicPlain || (PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletArabicDBPlain)
                text = number + ".";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanLCParenBoth)
                text = "(" + Settings.romanNumeralsLC[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanLCParenRight)
                text = Settings.romanNumeralsLC[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanLCPeriod)
                text = Settings.romanNumeralsLC[number - 1] + ".";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanUCParenBoth)
                text = "(" + Settings.romanNumeralsUC[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanUCParenRight)
                text = Settings.romanNumeralsUC[number - 1] + ")";
            else if((PpNumberedBulletStyle)style == PpNumberedBulletStyle.ppBulletRomanUCPeriod)
                text = Settings.romanNumeralsUC[number - 1] + ".";
            else
                text = number + ".";
            return text;
        }
    }
}

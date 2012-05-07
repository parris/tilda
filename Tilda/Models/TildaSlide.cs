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
using System.IO.Packaging;
using Ionic.Zip;
using System.Windows.Forms;

namespace Tilda.Models
{
    class TildaSlide
    {
        public PowerPoint.Slide slide;
        //public PowerPoint.Selection shapesRange;

        public TildaSlide(PowerPoint.Slide slide)
        {
            this.slide = slide;
        }

        public String exportSlide(){
            String js = "function(){";
            Dictionary<int, TildaShape> shapeMap = new Dictionary<int, TildaShape>(slide.Shapes.Count);
            List<PowerPoint.Shape> shapes = sortShapesByZIndex(slide.Shapes);

            //shapes, new count+1 for background
            int count = 0;


            foreach (PowerPoint.Shape shape in shapes) {
                if (shape.Type.Equals(Office.MsoShapeType.msoPlaceholder)||shape.Type.Equals(Office.MsoShapeType.msoTextBox)){
                    shapeMap.Add(shape.Id, new TildaTextbox(shape, count));
                } else //if (shape.Type.Equals(Office.MsoShapeType.msoPicture))
                    shapeMap.Add(shape.Id, new TildaPicture(shape, count)); //for now everything else can be an image!
                count++;
            }
            
            TildaAnimation[] animationMap = new TildaAnimation[slide.TimeLine.MainSequence.Count];
            int animationCount = 0;
            //animations started without click, on end, on start, etc
            foreach (PowerPoint.Effect effect in slide.TimeLine.MainSequence)
            {
                animationMap[animationCount] = new TildaAnimation(effect,shapeMap[effect.Shape.Id]);
                animationCount++;
            }

            js += "var idsToAnimate = new Array();";
            foreach (TildaShape shape in shapeMap.Values) {
                if (shape == null)
                    continue;
                js += shape.toRaphJS(animationMap);
            }

            js += this.exportBackgroundImage(shapes);
            js += "}";

            //js += .toRaphJS(shapeMap, animationMap);

            //animations via interaction, clicking
            /*foreach (PowerPoint.Sequence sequence in slide.TimeLine.InteractiveSequences){
                foreach (PowerPoint.Effect effect in sequence)
                {
                    Shape shape = effect.Shape;
                    float dur = effect.Timing.Duration;
                    //effect.
                }
            }*/
            return js;
        }

        private List<Shape> sortShapesByZIndex(PowerPoint.Shapes shapes){
            List<Shape> ordered = new List<Shape>();
            foreach(PowerPoint.Shape shape in shapes)
                ordered.Add(shape);

            ordered.Sort(new ZIndexShapeComparer());
            return ordered;
        }

        private String exportBackgroundImage(List<PowerPoint.Shape> shapes) {
            List<Shape> toBeUnhidden = new List<Shape>();

            foreach(PowerPoint.Shape shape in shapes) {
                //hide shapes that are not hidden
                if(shape.Visible == Office.MsoTriState.msoTrue) {
                    toBeUnhidden.Add(shape);
                    shape.Visible = Office.MsoTriState.msoFalse;
                }
            }

            String backgroundFileName = Settings.NextRandomValue() + "-" + Settings.NextRandomValue() + "-bg.png";
            String backgroundSavePath = Settings.outputMediaFullPath + Path.DirectorySeparatorChar + backgroundFileName;

            slide.Export(backgroundSavePath, "PNG",
                (int)(Settings.PresentationWidth() * 2), (int)(Settings.PresentationHeight() * 2));
            String js = "preso.shapes.push(preso.paper.image('" + Settings.outputMediaPath + "/" + backgroundFileName + "',0,0" + "," + (int)(Settings.PresentationWidth()) + "," + (int)(Settings.PresentationHeight()) + "));";
            js += "preso.shapes[(preso.shapes.length-1)].toBack();";

            //return shapes back to normal
            foreach(PowerPoint.Shape shape in toBeUnhidden)
                shape.Visible = Office.MsoTriState.msoTrue;

            return js;
        }

        private class ZIndexShapeComparer : IComparer<PowerPoint.Shape> {
            public int Compare(PowerPoint.Shape x, PowerPoint.Shape y) {
                if(x.ZOrderPosition > y.ZOrderPosition)
                    return 1;
                else if(x.ZOrderPosition < y.ZOrderPosition)
                    return -1;
                else
                    return 0;
            }
        }
    }
}

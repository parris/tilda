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
        public PowerPoint.Selection shapesRange;
        public int shapeCount = 0;

        public TildaSlide(PowerPoint.Slide slide)
        {
            this.slide = slide;
        }

        public String exportSlide(){
            String js = "function(){";

            //need to maintain id numbers of shapes. Shape id numbers include deleted shapes in indexing
            //We can settle for 5* the shape count for now. Too lazy to do it another way for now
            TildaShape[] shapeMap = new TildaShape[slide.Shapes.Count * 5]; 

            //shapes
            int count = 0;

            foreach (PowerPoint.Shape shape in slide.Shapes) {
                if (shape.Type.Equals(Office.MsoShapeType.msoPlaceholder)||shape.Type.Equals(Office.MsoShapeType.msoTextBox)){
                    shapeMap[shape.Id] = new TildaTextbox(shape,count);
                } else //if (shape.Type.Equals(Office.MsoShapeType.msoPicture))
                    shapeMap[shape.Id] = new TildaPicture(shape,count); //for now everything else can be an image!
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

            foreach (TildaShape shape in shapeMap) {
                if (shape == null)
                    continue;
                js += shape.toRaphJS(animationMap,this);
            }

            shapeCount = 0;
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

            //html += "</div>";
            return js;
        }
    }
}

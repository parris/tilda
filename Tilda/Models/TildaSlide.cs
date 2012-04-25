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

namespace Tilda.Models
{
    class TildaSlide
    {
        public PowerPoint.Slide slide;

        public TildaSlide(PowerPoint.Slide slide)
        {
            this.slide = slide;
        }

        public void saveSlideToLocation(String location)
        {
            String html = exportSlide(location);
        }

        /**
         * 
         */
        public String exportSlide(String location = ""){
            String js = "window.shapes = new Array();window.animations = new Array();";

            //sort of like a hash, not sure what the ids for the shapes will be but they definately wont be more than 3x the number of shapes
            TildaShape[] shapeMap = new TildaShape[slide.Shapes.Count * 3]; 
            //shapes
            int count = 0;

            foreach (PowerPoint.Shape shape in slide.Shapes) {
                shapeMap[shape.Id] = new TildaTextbox(shape,count);
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

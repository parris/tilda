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

    class TildaPicture : TildaShape {

        /**
         * Creates a new TildaPicture Object from a Powerpoint Shape that is actually an image
         * @param PowerPoint.Shape
         */
        public TildaPicture(PowerPoint.Shape shape, int id = 0)
            : base(shape, id) {
        }

        public override string toRaphJS(TildaAnimation[] animationMap, TildaSlide slide) {
            String fileName = (new Random(Int32.MaxValue)).Next().ToString() + "-image.png";
            String savePath = Settings.outputMediaFullPath + Path.DirectorySeparatorChar + fileName;

            this.shape.Export(savePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                (int)Settings.PresentationWidth() * 2, (int)Settings.PresentationHeight() * 2, PowerPoint.PpExportMode.ppScaleToFit);//widht&height*2 to support up 2x the size
            String js = "preso.shapes.push(preso.paper.image('" + Settings.outputMediaPath + "/" + fileName + "'," + this.position() + "," + shape.Width * scaler + "," + shape.Height * scaler + "));";
            foreach (TildaAnimation animation in animationMap)
                if (this.shape.Id.Equals(animation.shape.shape.Id)) {
                    js += "preso.shapes[" + slide.shapeCount + "].attr({'opacity':0});";
                    js += "preso.animations.push({'ids':[" + slide.shapeCount + "],'dur':" + animation.effect.Timing.Duration * 1000 + ",'delay':" + animation.effect.Timing.TriggerDelayTime * 1000 + ",animate:{'opacity':1}});";
                }
            slide.shapeCount++;
            return js;
        }
    }
}

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
            String savePath = "assets" + Path.DirectorySeparatorChar + fileName;
            this.shape.Export(slide.savePath + Path.DirectorySeparatorChar + savePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 
                slide.width*2,slide.height*2,PowerPoint.PpExportMode.ppScaleToFit);//widht&height*2 to support up 2x the size
            slide.shapeCount++;
            return "shapes.push(paper.image('" + "assets/" +  fileName + "'," + this.position() + "," + shape.Width * scaler + "," + shape.Height * scaler + "));";
        }
    }
}

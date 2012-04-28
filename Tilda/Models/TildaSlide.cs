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
        public int shapeCount = 0;
        public string savePath = "";

        //these are specified by powerpoint, sorta
        //TODO: allow to be specified by ppt's page setup, but then adjust raphael's paper obj to them as well
        public int width = 1024; 
        public int height = 768; 

        public TildaSlide(PowerPoint.Slide slide)
        {
            this.slide = slide;
            savePath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)+Path.DirectorySeparatorChar+"temp";
        }

        /**
         * 
         */
        public String exportSlide(){
            Directory.CreateDirectory(savePath);
            Directory.CreateDirectory(savePath + Path.DirectorySeparatorChar + "assets");
            String js = "function runSlide(){";

            //sort of like a hash, not sure what the ids for the shapes will be but they definately wont be more than 3x the number of shapes
            TildaShape[] shapeMap = new TildaShape[slide.Shapes.Count * 3]; 
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

            System.IO.File.WriteAllText(this.savePath+Path.DirectorySeparatorChar+"slide.js",js);
            //zip it all up
            using (ZipFile zip = new ZipFile()) {
                //make zip file
                //add content
                zip.AddFile(this.savePath + Path.DirectorySeparatorChar + "slide.js","");
                zip.AddDirectory(this.savePath + Path.DirectorySeparatorChar + "assets","assets");

                //add libs
                zip.AddDirectory(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "Web" + Path.DirectorySeparatorChar + "js", "js");
                zip.AddFile(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "Web" + Path.DirectorySeparatorChar + "index.html", "");

                zip.Save(savePath+Path.DirectorySeparatorChar+"slide.zip");
                //create dialog box
                SaveFileDialog dia = new SaveFileDialog();
                dia.Filter ="Zip File (*.zip)|*.zip|All files (*.*)|*.*";
                dia.FilterIndex = 2;
                dia.RestoreDirectory = true;

                //open dialog move file
                if(dia.ShowDialog() == DialogResult.OK)
                    File.Copy(savePath+Path.DirectorySeparatorChar+"slide.zip",dia.FileName,true);

                //remove zip file
                File.Delete(savePath + Path.DirectorySeparatorChar + "slide.zip");
            }

            Directory.Delete(savePath,true);

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
            return "Saved! :)";
        }
    }
}

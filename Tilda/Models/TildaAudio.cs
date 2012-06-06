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
    class TildaAudio : TildaShape {
        
        public TildaAudio(PowerPoint.Shape shape, int id = 0)
            : base(shape, id) {
        }

        /**
         * You can only export linked assets. It is way too much of pain to play embedded resources in PPT.
         */
        public override string toRaphJS() {
            String fileName = Settings.NextRandomValue() + "-" + Settings.NextRandomValue() + "-audio.mp3";
            String savePath = Settings.outputMediaFullPath + Path.DirectorySeparatorChar + fileName;
            if(this.shape.MediaFormat.IsLinked) {
                String source = this.shape.LinkFormat.SourceFullName;
                File.Copy(source, savePath);
                String output = "$('#audio-player').html(\"\");";
                output += "$('#audio-player').jPlayer({ready: function (event) {$(this).jPlayer(\"setMedia\","+
                    "{mp3:\"" + Settings.outputMediaPath + "/" + fileName + "\"}).jPlayer(\"play\");},"+
                    "solution:\"flash,html\",swfPath: \"\",supplied: \"mp3\",wmode: \"window\",autoplay:true});";
                return output;
            } else {
                return "";
            }
        }
    }
}
